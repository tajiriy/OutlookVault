Option Explicit On
Option Strict On
Option Infer Off

Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace Services

    ' ════════════════════════════════════════════════════════════
    '  進捗・結果データクラス
    ' ════════════════════════════════════════════════════════════

    ''' <summary>取り込み進捗の報告データ。IProgress(Of ImportProgress) で UI スレッドに通知する。</summary>
    Public Class ImportProgress
        Public Property FolderName As String
        Public Property ProcessedCount As Integer
        Public Property TotalCount As Integer
        Public Property CurrentSubject As String
    End Class

    ''' <summary>取り込み結果サマリー。</summary>
    Public Class ImportResult
        Public Property ImportedCount As Integer
        Public Property SkippedCount As Integer
        Public Property ErrorCount As Integer
        Public Property Errors As List(Of String)

        Public Sub New()
            Errors = New List(Of String)()
        End Sub
    End Class

    ' ════════════════════════════════════════════════════════════
    '  ImportService
    ' ════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Outlook からのメール取り込みを統括するサービスクラス。
    ''' OutlookService・ThreadingService・EmailRepository を組み合わせて
    ''' 重複排除・スレッド付与・DB 保存・添付ファイル保存を一括処理する。
    ''' </summary>
    Public Class ImportService

        Private ReadOnly _outlookSvc As OutlookService
        Private ReadOnly _repo As Data.EmailRepository
        Private ReadOnly _threadingSvc As ThreadingService
        Private ReadOnly _settings As Config.AppSettings

        Public Sub New(outlookSvc As OutlookService,
                       repo As Data.EmailRepository,
                       threadingSvc As ThreadingService,
                       settings As Config.AppSettings)
            _outlookSvc = outlookSvc
            _repo = repo
            _threadingSvc = threadingSvc
            _settings = settings
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  メインメソッド
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 指定フォルダのメールを取り込む。
        ''' 新着メールを最優先で取り込むため、受信日時の降順（最新から）でイテレートする。
        ''' </summary>
        ''' <param name="folderName">取り込み対象フォルダの表示名</param>
        ''' <param name="maxCount">今回の最大取り込み件数</param>
        ''' <param name="progress">進捗通知（Nothing の場合は通知しない）</param>
        Public Function ImportFolder(folderName As String,
                                     maxCount As Integer,
                                     Optional progress As IProgress(Of ImportProgress) = Nothing) As ImportResult
            Dim result As New ImportResult()

            Dim folder As Outlook.MAPIFolder = _outlookSvc.FindFolder(folderName)
            If folder Is Nothing Then
                result.Errors.Add("フォルダが見つかりません: " & folderName)
                result.ErrorCount += 1
                Return result
            End If

            Dim items As Outlook.Items = folder.Items
            Dim totalCount As Integer = items.Count
            Dim attachBaseDir As String = _settings.AttachmentDirectory

            ' 添付ファイル保存先ディレクトリを作成
            If Not IO.Directory.Exists(attachBaseDir) Then
                IO.Directory.CreateDirectory(attachBaseDir)
            End If

            ' 最新メールから処理するため末尾からイテレート
            Dim i As Integer = totalCount
            Do While i >= 1 AndAlso result.ImportedCount < maxCount
                Dim rawItem As Object = items.Item(i)
                i -= 1

                ' MailItem 以外（予定・タスク等）はスキップ
                If Not TypeOf rawItem Is Outlook.MailItem Then Continue Do

                Dim mailItem As Outlook.MailItem = CType(rawItem, Outlook.MailItem)
                Dim subject As String = mailItem.Subject

                Try
                    Dim imported As Boolean = ProcessMailItem(mailItem, attachBaseDir, result)
                    If imported Then
                        result.ImportedCount += 1
                    Else
                        result.SkippedCount += 1
                    End If
                Catch ex As Exception
                    result.ErrorCount += 1
                    result.Errors.Add("エラー [" & subject & "]: " & ex.Message)
                End Try

                ' 進捗通知
                If progress IsNot Nothing Then
                    Dim prog As New ImportProgress()
                    prog.FolderName = folderName
                    prog.ProcessedCount = result.ImportedCount + result.SkippedCount + result.ErrorCount
                    prog.TotalCount = totalCount
                    prog.CurrentSubject = subject
                    progress.Report(prog)
                End If
            Loop

            Return result
        End Function

        ''' <summary>複数フォルダをまとめて取り込む。</summary>
        Public Function ImportFolders(folderNames As IEnumerable(Of String),
                                      maxCountPerFolder As Integer,
                                      Optional progress As IProgress(Of ImportProgress) = Nothing) As ImportResult
            Dim total As New ImportResult()
            For Each name As String In folderNames
                Dim r As ImportResult = ImportFolder(name, maxCountPerFolder, progress)
                total.ImportedCount += r.ImportedCount
                total.SkippedCount += r.SkippedCount
                total.ErrorCount += r.ErrorCount
                total.Errors.AddRange(r.Errors)
            Next
            Return total
        End Function

        ' ════════════════════════════════════════════════════════════
        '  1件処理
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 1件の MailItem を処理する。
        ''' 重複・削除済みの場合は False を返す。新規保存した場合は True を返す。
        ''' </summary>
        Private Function ProcessMailItem(mailItem As Outlook.MailItem,
                                         attachBaseDir As String,
                                         importResult As ImportResult) As Boolean
            ' メールデータを抽出
            Dim email As Models.Email = _outlookSvc.ExtractEmailData(mailItem)

            ' 重複チェック（取り込み済み）
            If Not String.IsNullOrEmpty(email.MessageId) Then
                If _repo.MessageIdExists(email.MessageId) Then Return False
                If _repo.IsMessageIdDeleted(email.MessageId) Then Return False
            End If

            ' スレッド ID を付与
            _threadingSvc.AssignThreadId(email)

            ' DB に挿入
            Dim emailId As Integer = _repo.InsertEmail(email)

            ' 添付ファイルを保存して DB に登録
            If email.HasAttachments Then
                Dim attachSaveErrors As New List(Of String)()
                Dim attachments As List(Of Models.Attachment) =
                    _outlookSvc.SaveAttachments(mailItem, emailId, attachBaseDir, attachSaveErrors)
                For Each att As Models.Attachment In attachments
                    _repo.InsertAttachment(att)
                Next
                For Each errMsg As String In attachSaveErrors
                    importResult.Errors.Add(errMsg)
                    importResult.ErrorCount += 1
                Next
            End If

            Return True
        End Function

    End Class

End Namespace
