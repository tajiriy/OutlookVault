Option Explicit On
Option Strict On
Option Infer Off

Imports System.Threading
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

    ''' <summary>取り込みエラー1件の詳細情報。</summary>
    Public Class ImportErrorEntry
        Public Property Timestamp As DateTime
        Public Property FolderName As String
        Public Property MessageId As String
        Public Property Subject As String
        Public Property ErrorMessage As String

        Public Sub New(folderName As String, messageId As String, subject As String, errorMessage As String)
            Me.Timestamp = DateTime.Now
            Me.FolderName = folderName
            Me.MessageId = messageId
            Me.Subject = subject
            Me.ErrorMessage = errorMessage
        End Sub
    End Class

    ''' <summary>取り込み結果サマリー。</summary>
    Public Class ImportResult
        Public Property ImportedCount As Integer
        Public Property SkippedCount As Integer
        Public Property ErrorCount As Integer
        Public Property TotalOutlookCount As Integer
        Public Property DeletedCount As Integer
        Public Property Errors As List(Of ImportErrorEntry)

        Public Sub New()
            Errors = New List(Of ImportErrorEntry)()
        End Sub
    End Class

    ''' <summary>削除同期の進捗報告データ。</summary>
    Public Class SyncDeletionProgress
        Public Property FolderName As String
        Public Property ScannedCount As Integer
        Public Property TotalCount As Integer
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
                                     Optional progress As IProgress(Of ImportProgress) = Nothing,
                                     Optional cancellationToken As CancellationToken = Nothing) As ImportResult
            Dim result As New ImportResult()

            Dim folder As Outlook.MAPIFolder = _outlookSvc.FindFolder(folderName)
            If folder Is Nothing Then
                Logger.Warn("フォルダが見つかりません: " & folderName)
                result.Errors.Add(New ImportErrorEntry(folderName, "", "", "フォルダが見つかりません: " & folderName))
                result.ErrorCount += 1
                Return result
            End If

            Dim items As Outlook.Items = folder.Items
            Dim totalCount As Integer = items.Count
            Logger.Info(String.Format("フォルダ '{0}' の取り込みを開始します（{1}件）", folderName, totalCount))
            result.TotalOutlookCount += totalCount
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

            ' ── 高速インポート: synchronous OFF ──
            Dim perfConn As System.Data.SQLite.SQLiteConnection = _dbManager.GetConnection()
            Try
                _dbManager.SetSynchronousMode(perfConn, "OFF")
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
                    cancellationToken.ThrowIfCancellationRequested()
                    Dim rawItem As Object = items.Item(i)
                    i += stepDir

                    ' MailItem 以外（予定・タスク等）はスキップ
                    If Not TypeOf rawItem Is Outlook.MailItem Then Continue Do

                    Dim mailItem As Outlook.MailItem = CType(rawItem, Outlook.MailItem)
                    Dim subject As String = mailItem.Subject

                    Try
                        Dim imported As Boolean = ProcessMailItem(mailItem, attachBaseDir, existingIds, deletedIds, result, folderName)
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
                        Dim errMsgId As String = ""
                        Try
                            errMsgId = _outlookSvc.ExtractMessageId(mailItem)
                        Catch
                        End Try
                        result.Errors.Add(New ImportErrorEntry(folderName, errMsgId, subject, ex.Message))
                        Logger.Error(String.Format("メール取り込みエラー — フォルダ: {0}, 件名: {1}", folderName, subject), ex)
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

                Logger.Info(String.Format("フォルダ '{0}' の取り込みが完了しました — 新規: {1}件, スキップ: {2}件, エラー: {3}件",
                    folderName, result.ImportedCount, result.SkippedCount, result.ErrorCount))

            Catch ex As OperationCanceledException
                _repo.RollbackBulk()
                Logger.Info(String.Format("フォルダ '{0}' の取り込みが中断されました", folderName))
                Throw
            Catch ex As Exception
                _repo.RollbackBulk()
                Logger.Error(String.Format("フォルダ '{0}' の取り込み中にエラーが発生しました", folderName), ex)
                Throw
            Finally
                _threadingSvc.ClearCaches()
                ' ── 高速インポート後処理: synchronous 復元 ──
                Try
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
                                      Optional progress As IProgress(Of ImportProgress) = Nothing,
                                      Optional cancellationToken As CancellationToken = Nothing) As ImportResult
            Dim total As New ImportResult()

            ' 全 MessageID を一括キャッシュ（1件ずつ SQL を発行するより大幅に高速）
            Dim existingIds As HashSet(Of String) = _repo.GetAllMessageIds()
            Dim deletedIds As HashSet(Of String) = _repo.GetAllDeletedMessageIds()

            ' Exchange アドレスキャッシュを DB からロード
            _outlookSvc.LoadExchangeCache(_repo.LoadExchangeAddressCache())

            For Each name As String In folderNames
                cancellationToken.ThrowIfCancellationRequested()
                Dim r As ImportResult = ImportFolder(name, maxCountPerFolder, existingIds, deletedIds, progress, cancellationToken)
                total.ImportedCount += r.ImportedCount
                total.SkippedCount += r.SkippedCount
                total.ErrorCount += r.ErrorCount
                total.TotalOutlookCount += r.TotalOutlookCount
                total.Errors.AddRange(r.Errors)
            Next

            ' Exchange アドレスキャッシュの新規分を DB に書き戻し
            _repo.SaveExchangeAddressCache(_outlookSvc.GetExchangeCache())

            Return total
        End Function

        ' ════════════════════════════════════════════════════════════
        '  削除同期
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 指定フォルダで Outlook 側から削除されたメールをアーカイブ DB からも削除する。
        ''' Outlook フォルダ内の全 MailItem の MessageID を取得し、DB 側と突合して
        ''' Outlook に存在しないものを削除する。
        ''' </summary>
        Public Function SyncDeletions(folderName As String,
                                       Optional progress As IProgress(Of SyncDeletionProgress) = Nothing,
                                       Optional cancellationToken As CancellationToken = Nothing) As Integer
            Dim folder As Outlook.MAPIFolder = _outlookSvc.FindFolder(folderName)
            If folder Is Nothing Then
                Logger.Warn("削除同期: フォルダが見つかりません: " & folderName)
                Return 0
            End If

            Logger.Info(String.Format("フォルダ '{0}' の削除同期を開始します", folderName))

            ' Outlook 側の全 MessageID を取得（軽量スキャン）
            Dim scanProgress As IProgress(Of Integer) = Nothing
            If progress IsNot Nothing Then
                Dim totalItems As Integer = folder.Items.Count
                Dim fname As String = folderName
                scanProgress = New Progress(Of Integer)(
                    Sub(scanned As Integer)
                        Dim p As New SyncDeletionProgress()
                        p.FolderName = fname
                        p.ScannedCount = scanned
                        p.TotalCount = totalItems
                        progress.Report(p)
                    End Sub)
            End If

            Dim outlookIds As HashSet(Of String) = _outlookSvc.GetFolderMessageIds(folder, scanProgress)

            ' DB 側のフォルダ内メールを取得
            Dim archivedEmails As Dictionary(Of String, Integer) = _repo.GetMessageIdsByFolder(folderName)

            ' Outlook に存在しないメールを特定
            Dim idsToDelete As New List(Of Integer)()
            For Each kvp As KeyValuePair(Of String, Integer) In archivedEmails
                If Not outlookIds.Contains(kvp.Key) Then
                    idsToDelete.Add(kvp.Value)
                End If
            Next

            If idsToDelete.Count = 0 Then
                Logger.Info(String.Format("フォルダ '{0}' の削除同期が完了しました — 削除対象なし", folderName))
                Return 0
            End If

            ' 一括削除（添付ファイルパスも取得）
            Dim attachmentPaths As List(Of String) = _repo.DeleteEmailsByIds(idsToDelete)

            ' 添付ファイルの物理削除
            For Each filePath As String In attachmentPaths
                Try
                    If IO.File.Exists(filePath) Then
                        IO.File.Delete(filePath)
                    End If
                Catch
                    ' 物理削除失敗は無視
                End Try
            Next

            Logger.Info(String.Format("フォルダ '{0}' の削除同期が完了しました — {1}件削除", folderName, idsToDelete.Count))
            Return idsToDelete.Count
        End Function

        ''' <summary>複数フォルダの削除同期を実行する。</summary>
        Public Function SyncDeletionsForFolders(folderNames As IEnumerable(Of String),
                                                 Optional progress As IProgress(Of SyncDeletionProgress) = Nothing,
                                                 Optional cancellationToken As CancellationToken = Nothing) As Integer
            Dim totalDeleted As Integer = 0
            For Each name As String In folderNames
                cancellationToken.ThrowIfCancellationRequested()
                totalDeleted += SyncDeletions(name, progress, cancellationToken)
            Next
            Return totalDeleted
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
                                         importResult As ImportResult,
                                         folderName As String) As Boolean
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
                    importResult.Errors.Add(New ImportErrorEntry(folderName, email.MessageId, email.Subject, errMsg))
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
