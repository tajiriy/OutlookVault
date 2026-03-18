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
        Private ReadOnly _dbManager As Data.DatabaseManager

        Public Sub New(outlookSvc As OutlookService,
                       repo As Data.EmailRepository,
                       threadingSvc As ThreadingService,
                       settings As Config.AppSettings,
                       dbManager As Data.DatabaseManager)
            _outlookSvc = outlookSvc
            _repo = repo
            _threadingSvc = threadingSvc
            _settings = settings
            _dbManager = dbManager
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  メインメソッド
        ' ════════════════════════════════════════════════════════════

        ''' <summary>バルクトランザクションを N 件ごとに中間コミットする件数。</summary>
        Private Const BulkCommitInterval As Integer = 200

        ''' <summary>
        ''' 指定フォルダのメールを取り込む。
        ''' 設定の ImportOldestFirst に応じて古い順または新しい順でイテレートする。
        ''' </summary>
        ''' <param name="folderName">取り込み対象フォルダの表示名</param>
        ''' <param name="maxCount">今回の最大取り込み件数</param>
        ''' <param name="existingIds">取り込み済み MessageID のキャッシュ</param>
        ''' <param name="deletedIds">削除済み MessageID のキャッシュ</param>
        ''' <param name="progress">進捗通知（Nothing の場合は通知しない）</param>

        Public Function ImportFolder(folderName As String,
                                     maxCount As Integer,
                                     existingIds As HashSet(Of String),
                                     deletedIds As HashSet(Of String),
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

            ' ── スレッド判定インメモリキャッシュをロード ─────────────
            Dim messageIdMap As Dictionary(Of String, String) = Nothing
            Dim subjectMap As Dictionary(Of String, String) = Nothing
            _repo.GetThreadIdCaches(messageIdMap, subjectMap)
            _threadingSvc.LoadCaches(messageIdMap, subjectMap)

            ' ── 高速インポート: synchronous OFF + FTS トリガー無効化 ──
            Dim perfConn As System.Data.SQLite.SQLiteConnection = _dbManager.GetConnection()
            Try
                _dbManager.SetSynchronousMode(perfConn, "OFF")
                _repo.DisableFtsTriggers(perfConn)
            Catch
                perfConn.Dispose()
                Throw
            End Try

            ' ── DB バルク書き込みモード開始 ───────────────────────────
            _repo.BeginBulk()
            Dim sinceLastCommit As Integer = 0

            Try
                ' 設定に応じて古い順（1→N）または新しい順（N→1）でイテレート
                Dim oldestFirst As Boolean = _settings.ImportOldestFirst
                Dim i As Integer = If(oldestFirst, 1, totalCount)
                Dim stepDir As Integer = If(oldestFirst, 1, -1)
                Dim endCond As Func(Of Boolean) = If(oldestFirst,
                    Function() i <= totalCount,
                    Function() i >= 1)
                Do While endCond() AndAlso result.ImportedCount < maxCount
                    Dim rawItem As Object = items.Item(i)
                    i += stepDir

                    ' MailItem 以外（予定・タスク等）はスキップ
                    If Not TypeOf rawItem Is Outlook.MailItem Then Continue Do

                    Dim mailItem As Outlook.MailItem = CType(rawItem, Outlook.MailItem)
                    Dim subject As String = mailItem.Subject

                    Try
                        Dim imported As Boolean = ProcessMailItem(mailItem, attachBaseDir, existingIds, deletedIds, result)
                        If imported Then
                            result.ImportedCount += 1
                            sinceLastCommit += 1
                            ' N 件ごとに中間コミット（クラッシュ時のデータロスを軽減）
                            If sinceLastCommit >= BulkCommitInterval Then
                                _repo.CommitBulk()
                                _repo.BeginBulk()
                                sinceLastCommit = 0
                            End If
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

                ' 残分をコミット
                _repo.CommitBulk()

            Catch
                _repo.RollbackBulk()
                Throw
            Finally
                _threadingSvc.ClearCaches()
                ' ── 高速インポート後処理: FTS 再構築 + synchronous 復元 ──
                Try
                    _repo.RebuildFtsIndex(perfConn)
                    _repo.EnableFtsTriggers(perfConn)
                    _dbManager.SetSynchronousMode(perfConn, "NORMAL")
                Finally
                    perfConn.Dispose()
                End Try
            End Try

            Return result
        End Function

        ''' <summary>複数フォルダをまとめて取り込む。</summary>
        Public Function ImportFolders(folderNames As IEnumerable(Of String),
                                      maxCountPerFolder As Integer,
                                      Optional progress As IProgress(Of ImportProgress) = Nothing) As ImportResult
            Dim total As New ImportResult()

            ' 全 MessageID を一括キャッシュ（1件ずつ SQL を発行するより大幅に高速）
            Dim existingIds As HashSet(Of String) = _repo.GetAllMessageIds()
            Dim deletedIds As HashSet(Of String) = _repo.GetAllDeletedMessageIds()

            ' Exchange アドレスキャッシュを DB からロード
            _outlookSvc.LoadExchangeCache(_repo.LoadExchangeAddressCache())

            For Each name As String In folderNames
                Dim r As ImportResult = ImportFolder(name, maxCountPerFolder, existingIds, deletedIds, progress)
                total.ImportedCount += r.ImportedCount
                total.SkippedCount += r.SkippedCount
                total.ErrorCount += r.ErrorCount
                total.Errors.AddRange(r.Errors)
            Next

            ' Exchange アドレスキャッシュの新規分を DB に書き戻し
            _repo.SaveExchangeAddressCache(_outlookSvc.GetExchangeCache())

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
                                         existingIds As HashSet(Of String),
                                         deletedIds As HashSet(Of String),
                                         importResult As ImportResult) As Boolean
            ' 軽量に MessageID だけ取得して重複チェック（本文・受信者等の重いCOM操作を回避）
            Dim messageId As String = _outlookSvc.ExtractMessageId(mailItem)
            If Not String.IsNullOrEmpty(messageId) Then
                If existingIds.Contains(messageId) Then Return False
                If deletedIds.Contains(messageId) Then Return False
            End If

            ' 新規メールのみフルデータを抽出
            Dim email As Models.Email = _outlookSvc.ExtractEmailData(mailItem)

            ' スレッド ID を付与
            _threadingSvc.AssignThreadId(email)

            ' DB に挿入
            Dim emailId As Integer = _repo.InsertEmail(email)

            ' キャッシュに追加（同一セッション内の重複防止）
            If Not String.IsNullOrEmpty(email.MessageId) Then
                existingIds.Add(email.MessageId)
            End If

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

                ' インライン画像のみで実質的な添付がない場合はフラグを修正
                Dim hasRealAttachment As Boolean = False
                For Each att As Models.Attachment In attachments
                    If Not att.IsInline Then
                        hasRealAttachment = True
                        Exit For
                    End If
                Next
                If Not hasRealAttachment Then
                    _repo.UpdateHasAttachments(emailId, False)
                End If
            End If

            Return True
        End Function

    End Class

End Namespace
