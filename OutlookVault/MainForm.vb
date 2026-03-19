Option Explicit On
Option Strict On
Option Infer Off

Imports System.Runtime.InteropServices
Imports System.Threading.Tasks

''' <summary>メール一覧 ListView の列インデックス。</summary>
Public Enum EmailListColumn As Integer
    Attachment = 0
    Subject = 1
    Sender = 2
    ReceivedAt = 3
    Size = 4
End Enum

Public Class MainForm

    ' ════════════════════════════════════════════════════════════
    '  Win32 API (ListView 全選択用)
    ' ════════════════════════════════════════════════════════════
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByRef lParam As LVITEM) As IntPtr
    End Function

    Private Const LVM_FIRST As Integer = &H1000
    Private Const LVM_SETITEMSTATE As Integer = LVM_FIRST + 43

    <StructLayout(LayoutKind.Sequential)>
    Private Structure LVITEM
        Public mask As Integer
        Public iItem As Integer
        Public iSubItem As Integer
        Public state As Integer
        Public stateMask As Integer
        Public pszText As IntPtr
        Public cchTextMax As Integer
        Public iImage As Integer
        Public lParam As IntPtr
    End Structure

    Private Const LVIS_SELECTED As Integer = &H2
    Private Const LVIS_FOCUSED As Integer = &H1
    Private Const LVIF_STATE As Integer = &H8

    ' ════════════════════════════════════════════════════════════
    '  フィールド
    ' ════════════════════════════════════════════════════════════

    Private _settings As Config.AppSettings
    Private _dbManager As Data.DatabaseManager
    Private _repo As Data.EmailRepository
    ' _autoImportTimer, _scheduledImportTimer は Designer.vb で宣言・生成
    Private _lastScheduledRunDate As DateTime = DateTime.MinValue

    Private Const TrashFolderTag As String = "__TRASH__"  ' ゴミ箱ノードの Tag 識別子

    Private _emailCache As List(Of Models.Email)
    Private _currentFolder As String      ' Nothing = すべて
    Private _isTrashView As Boolean       ' True = ゴミ箱表示中
    Private _isImporting As Boolean
    Private _importCts As System.Threading.CancellationTokenSource
    Private _lastOutlookTotalCount As Integer = -1  ' -1 = 未取得
    Private _autoImportEnabled As Boolean
    Private _isRealClose As Boolean       ' True = 本当に終了する（トレイ格納ではなく）
    Private _searchQuery As String        ' 現在の検索クエリ（Nothing = 検索なし）
    Private _openEmailWindows As New Dictionary(Of Integer, Forms.EmailViewForm)() ' メールID → 開いているビューウィンドウ
    Private _updatingFolderCounts As Boolean ' フォルダ件数更新中のイベント抑制フラグ
    Private _currentFolderTotalCount As Integer ' 現在選択中フォルダの総件数
    Private _loadVersion As Integer             ' LoadEmailsAsync の競合検出用バージョン番号

    ' 列インデックス: 0=添付 1=件名 2=差出人 3=受信日時 4=サイズ
    Private _sortColumn As EmailListColumn = EmailListColumn.ReceivedAt
    Private _sortAscending As Boolean = False   ' デフォルト: 降順
    Private ReadOnly _colBaseNames As String() = {"添付", "件名", "差出人", "受信日時", "サイズ"}

    ' ════════════════════════════════════════════════════════════
    '  フォーム初期化
    ' ════════════════════════════════════════════════════════════

    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' アプリケーションアイコンを設定
        Dim icoPath As String = System.IO.Path.Combine(Application.StartupPath, "app.ico")
        If System.IO.File.Exists(icoPath) Then
            Me.Icon = New Drawing.Icon(icoPath)
        End If

        ' メール一覧高さ: コンテンツ領域の約 40%（splitMain はデザイナー値をそのまま使用）
        Dim contentHeight As Integer = Me.ClientSize.Height - menuStrip.Height - toolStrip.Height - statusStrip.Height
        splitRight.SplitterDistance = CInt(contentHeight * 0.4)

        Services.Logger.Info("アプリケーションを起動しました")

        InitializeServices()
        SetupAutoImportTimer()
        SetupNotifyIcon()
        SetupEmailListColumns()
        SetupToggleViewButton()
        UpdateContextMenuForView()
        LoadFolderTree()
        Await UpdateStatusBarAsync()

        ' ゴミ箱の自動パージ
        PurgeExpiredTrash()

        Services.Logger.Info("初期化が完了しました")

        If _settings.AutoImportEnabled Then
            StartAutoImport()
        End If

        ' --minimized 引数付きで起動された場合はトレイに直接格納
        Dim args() As String = Environment.GetCommandLineArgs()
        For Each arg As String In args
            If String.Equals(arg, "--minimized", StringComparison.OrdinalIgnoreCase) Then
                MinimizeToSystemTray()
                Services.Logger.Info("--minimized 引数によりトレイに格納しました")
                Exit For
            End If
        Next
    End Sub

    Private Sub InitializeServices()
        _settings = Config.AppSettings.Instance
        _dbManager = New Data.DatabaseManager(_settings.DbFilePath)
        _dbManager.Initialize()
        _repo = New Data.EmailRepository(_dbManager)
        _emailCache = New List(Of Models.Email)()
    End Sub

    ''' <summary>メール一覧の列設定を復元し、添付アイコンを登録する。</summary>
    ''' <remarks>列定義・ImageList・ContextMenu は Designer.vb で生成済み。</remarks>
    Private Sub SetupEmailListColumns()
        ' ── 添付アイコンを ImageList に登録 ───────────────────────
        emailImageList.Images.Add(New Drawing.Bitmap(16, 16))  ' index 0: 空白（添付なし）
        emailImageList.Images.Add(CreatePaperclipIcon())       ' index 1: ペーパークリップ

        ' ── ソート設定を復元 ──────────────────────────────────────
        _sortColumn = CType(_settings.EmailListSortColumn, EmailListColumn)
        _sortAscending = _settings.EmailListSortAscending

        ' ── フォルダ別列設定を復元（なければグローバル設定を使用） ──
        RestoreColumnSettingsForFolder(GetCurrentFolderKey())
    End Sub

    ''' <summary>16×16 のペーパークリップアイコンを GDI+ で描画して返す。</summary>
    Private Function CreatePaperclipIcon() As Drawing.Bitmap
        Dim bmp As New Drawing.Bitmap(16, 16)
        Using g As Drawing.Graphics = Drawing.Graphics.FromImage(bmp)
            g.Clear(Drawing.Color.Transparent)
            g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
            Using pen As New Drawing.Pen(Drawing.Color.FromArgb(80, 120, 190), 1.5F)
                pen.StartCap = Drawing.Drawing2D.LineCap.Round
                pen.EndCap = Drawing.Drawing2D.LineCap.Round
                g.DrawArc(pen, 2, 1, 10, 13, 40, 280)
                g.DrawArc(pen, 4, 4, 6, 7, 40, 180)
                g.DrawLine(pen, 12, 9, 12, 14)
            End Using
        End Using
        Return bmp
    End Function

    ''' <summary>メールサイズをKB/MB単位の表示文字列に変換する。</summary>
    Private Shared Function FormatEmailSize(sizeBytes As Long) As String
        If sizeBytes <= 0 Then Return String.Empty
        Return Services.FileHelper.FormatFileSize(sizeBytes)
    End Function

    ''' <summary>タイマーの間隔を設定から反映する。</summary>
    ''' <remarks>タイマーのインスタンス生成は Designer.vb で実施済み。</remarks>
    Private Sub SetupAutoImportTimer()
        _autoImportTimer.Interval = _settings.AutoImportIntervalMinutes * 60 * 1000
    End Sub

    ' SetupListViewContextMenu は不要（Designer.vb で設定済み）

    ' ════════════════════════════════════════════════════════════
    '  フォルダツリー
    ' ════════════════════════════════════════════════════════════

    Private Sub LoadFolderTree()
        treeViewFolders.BeginUpdate()
        treeViewFolders.Nodes.Clear()

        Dim folderData As Tuple(Of Integer, Dictionary(Of String, Integer)) = _repo.GetFolderCounts()
        Dim totalCount As Integer = folderData.Item1
        Dim folderCounts As Dictionary(Of String, Integer) = folderData.Item2

        Dim nodeAll As New TreeNode(String.Format("すべて ({0:N0})", totalCount))
        nodeAll.Tag = Nothing
        treeViewFolders.Nodes.Add(nodeAll)

        For Each kvp As KeyValuePair(Of String, Integer) In folderCounts
            Dim node As New TreeNode(String.Format("{0} ({1:N0})", kvp.Key, kvp.Value))
            node.Tag = kvp.Key
            treeViewFolders.Nodes.Add(node)
        Next

        ' ゴミ箱ノード
        Dim trashCount As Integer = _repo.GetTrashCount()
        Dim nodeTrash As New TreeNode(String.Format("ゴミ箱 ({0:N0})", trashCount))
        nodeTrash.Tag = TrashFolderTag
        treeViewFolders.Nodes.Add(nodeTrash)

        treeViewFolders.EndUpdate()
        treeViewFolders.SelectedNode = nodeAll
    End Sub

    ''' <summary>フォルダツリーの右クリックでゴミ箱ノードのコンテキストメニューを表示する。</summary>
    Private Sub treeViewFolders_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles treeViewFolders.NodeMouseClick
        If e.Button = MouseButtons.Right Then
            treeViewFolders.SelectedNode = e.Node
            Dim tag As String = TryCast(e.Node.Tag, String)
            If String.Equals(tag, TrashFolderTag, StringComparison.Ordinal) Then
                folderContextMenu.Show(treeViewFolders, e.Location)
            End If
        End If
    End Sub

    Private Async Sub treeViewFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterSelect
        If e.Node Is Nothing Then Return
        If _updatingFolderCounts Then Return

        ' 切り替え前のフォルダの列設定を保存
        SaveColumnSettingsForFolder(GetCurrentFolderKey())

        Dim tag As String = TryCast(e.Node.Tag, String)
        _isTrashView = String.Equals(tag, TrashFolderTag, StringComparison.Ordinal)

        If _isTrashView Then
            _currentFolder = Nothing
        Else
            _currentFolder = tag
        End If

        ' 検索ボックスと検索クエリをクリアしてフォルダ切り替え
        txtSearch.Text = String.Empty
        _searchQuery = Nothing
        UpdateContextMenuForView()

        ' 切り替え後のフォルダの列設定を復元
        RestoreColumnSettingsForFolder(GetCurrentFolderKey())

        Await LoadEmailsAsync(_currentFolder)
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  メール一覧（VirtualMode）
    ' ════════════════════════════════════════════════════════════

    ''' <summary>DBからメールを読み込んでListViewのVirtualListSizeを更新する（非同期）。</summary>
    ''' <remarks>
    ''' フォルダを素早く切り替えた場合に古い結果が新しい結果を上書きしないよう、
    ''' バージョン番号で最新のリクエストのみ UI に反映する。
    ''' </remarks>
    Private Async Function LoadEmailsAsync(Optional folderName As String = Nothing) As Task
        _loadVersion += 1
        Dim myVersion As Integer = _loadVersion
        Dim repo As Data.EmailRepository = _repo
        Dim isTrash As Boolean = _isTrashView
        Dim emails As List(Of Models.Email) = Await Task.Run(
            Function() As List(Of Models.Email)
                If isTrash Then
                    Return repo.GetTrashEmails()
                Else
                    Return repo.GetEmailsForList(folderName)
                End If
            End Function)

        ' 待機中に別のロードが開始された場合は結果を破棄する
        If myVersion <> _loadVersion Then Return

        _emailCache = emails
        _currentFolderTotalCount = _emailCache.Count
        SortEmailCache()
        listViewEmails.VirtualListSize = _emailCache.Count
        listViewEmails.Invalidate()
        UpdateFolderCountLabel()
        Await UpdateStatusBarAsync()
    End Function

    ''' <summary>メール一覧で行が選択されたときにプレビューと会話ビューを更新する。</summary>
    Private Sub listViewEmails_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listViewEmails.SelectedIndexChanged
        If listViewEmails.SelectedIndices.Count = 0 Then
            emailPreview.ClearPreview()
            conversationView.ClearView()
            UpdateToggleViewButton()
            Return
        End If

        Dim idx As Integer = listViewEmails.SelectedIndices(0)
        If idx < 0 OrElse idx >= _emailCache.Count Then Return

        Dim email As Models.Email = _repo.GetEmailById(_emailCache(idx).Id)
        If email IsNot Nothing Then
            emailPreview.ShowEmail(email, _searchQuery)
            LoadConversationView(email)
            UpdateToggleViewButton()
        End If
    End Sub

    ''' <summary>選択メールのスレッドを取得して会話ビューに表示する。</summary>
    Private Sub LoadConversationView(email As Models.Email)
        Dim threadEmails As List(Of Models.Email)
        If String.IsNullOrEmpty(email.ThreadId) Then
            ' スレッドIDなし → 単独メールとして表示
            threadEmails = New List(Of Models.Email)() From {email}
        Else
            threadEmails = _repo.GetEmailsByThreadId(email.ThreadId)
            If threadEmails.Count = 0 Then
                threadEmails = New List(Of Models.Email)() From {email}
            End If
        End If
        conversationView.ShowThread(threadEmails, email.Id)
    End Sub

    ''' <summary>メール一覧のダブルクリックで別ウィンドウにメールを表示する。</summary>
    Private Sub listViewEmails_DoubleClick(sender As Object, e As EventArgs) Handles listViewEmails.DoubleClick
        If listViewEmails.SelectedIndices.Count = 0 Then Return

        Dim idx As Integer = listViewEmails.SelectedIndices(0)
        If idx < 0 OrElse idx >= _emailCache.Count Then Return

        Dim emailId As Integer = _emailCache(idx).Id

        ' 既に同じメールのウィンドウが開いている場合は前面に表示
        If _openEmailWindows.ContainsKey(emailId) Then
            Dim existing As Forms.EmailViewForm = _openEmailWindows(emailId)
            If Not existing.IsDisposed Then
                If existing.WindowState = FormWindowState.Minimized Then
                    existing.WindowState = FormWindowState.Normal
                End If
                existing.Activate()
                Return
            Else
                _openEmailWindows.Remove(emailId)
            End If
        End If

        Dim email As Models.Email = _repo.GetEmailById(emailId)
        If email Is Nothing Then Return

        Dim frm As New Forms.EmailViewForm()
        frm.ShowEmail(email, _searchQuery)

        ' ウィンドウを辞書に登録し、閉じたときに削除する
        _openEmailWindows(emailId) = frm
        Dim capturedId As Integer = emailId
        Dim handler As FormClosedEventHandler = Nothing
        handler = Sub(s As Object, ev As FormClosedEventArgs)
                      _openEmailWindows.Remove(capturedId)
                      RemoveHandler frm.FormClosed, handler
                  End Sub
        AddHandler frm.FormClosed, handler
        frm.Show()
    End Sub

    ''' <summary>VirtualMode: 指定インデックスのListViewItemを返す。</summary>
    Private Sub listViewEmails_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs) Handles listViewEmails.RetrieveVirtualItem
        If e.ItemIndex < 0 OrElse e.ItemIndex >= _emailCache.Count Then
            e.Item = New ListViewItem(String.Empty)
            Return
        End If

        Dim email As Models.Email = _emailCache(e.ItemIndex)
        Dim subject As String = If(String.IsNullOrEmpty(email.Subject), "(件名なし)", email.Subject)
        Dim displaySender As String = If(Not String.IsNullOrEmpty(email.SenderName), email.SenderName, email.SenderEmail)

        ' col 0: 添付（SmallImageList アイコン、テキストなし）
        Dim item As New ListViewItem(String.Empty)
        item.ImageIndex = If(email.HasAttachments, 1, 0)
        item.SubItems.Add(subject)                                       ' col 1: 件名
        item.SubItems.Add(If(displaySender, String.Empty))               ' col 2: 差出人
        item.SubItems.Add(email.ReceivedAt.ToString("yyyy/MM/dd HH:mm")) ' col 3: 受信日時
        item.SubItems.Add(FormatEmailSize(email.EmailSize))              ' col 4: サイズ
        ' ゴミ箱表示時: 追加列
        If _isTrashView Then
            item.SubItems.Add(If(email.FolderName, String.Empty))                            ' col 5: フォルダ
            If email.DeletedAt.HasValue Then
                item.SubItems.Add(email.DeletedAt.Value.ToString("yyyy/MM/dd HH:mm"))        ' col 6: 削除日時
                Dim trashDays As Integer = _settings.TrashAutoDeleteDays
                If trashDays > 0 Then
                    item.SubItems.Add(email.DeletedAt.Value.AddDays(trashDays).ToString("yyyy/MM/dd HH:mm")) ' col 7: 削除予定日時
                Else
                    item.SubItems.Add("（自動削除なし）")
                End If
            Else
                item.SubItems.Add(String.Empty)
                item.SubItems.Add(String.Empty)
            End If
        End If
        item.Tag = email.Id
        e.Item = item
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  取り込み
    ' ════════════════════════════════════════════════════════════

    Private Async Sub btnImportNow_Click(sender As Object, e As EventArgs) Handles btnImportNow.Click, menuItemImportNow.Click
        Await RunImportAsync()
    End Sub

    Private Sub menuItemImportCancel_Click(sender As Object, e As EventArgs) Handles menuItemImportCancel.Click
        If Not _isImporting OrElse _importCts Is Nothing Then Return

        Dim answer As DialogResult = MessageBox.Show(
            "取り込みを中止しますか？" & vbCrLf &
            "（処理中のバッチはロールバックされます）",
            "取り込み中止の確認",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2)

        If answer = DialogResult.Yes Then
            _importCts.Cancel()
        End If
    End Sub

    ''' <summary>
    ''' 取り込みを実行する。fullSync=True の場合は最大件数制限なしで完全同期する。
    ''' Outlook COM は STA スレッドが必要なため、TaskCompletionSource + 専用STAスレッドで実行する。
    ''' </summary>
    Private Async Function RunImportAsync(Optional fullSync As Boolean = False) As Task
        If _isImporting Then Return
        _isImporting = True
        _importCts = New System.Threading.CancellationTokenSource()
        btnImportNow.Enabled = False
        menuItemImportNow.Enabled = False
        menuItemImportCancel.Enabled = True
        lblStatusCount.Text = "取り込み中..."

        Services.Logger.Info("取り込みを開始します")
        Dim importStopwatch As New System.Diagnostics.Stopwatch()
        importStopwatch.Start()

        Try
            Dim progress As New Progress(Of Services.ImportProgress)(
                Sub(p)
                    lblStatusCount.Text = String.Format("取り込み中... {0}/{1} ({2})",
                        p.ProcessedCount, p.TotalCount, p.FolderName)
                End Sub)

            Dim tcs As New TaskCompletionSource(Of Services.ImportResult)()
            Dim targetFolders As List(Of String) = _settings.TargetFolders
            Dim maxCount As Integer = If(fullSync, Integer.MaxValue, _settings.MaxImportCount)
            Dim repo As Data.EmailRepository = _repo
            Dim settings As Config.AppSettings = _settings
            Dim dbManager As Data.DatabaseManager = _dbManager
            Dim syncDeletions As Boolean = _settings.SyncDeletionsEnabled

            Dim syncProgress As Progress(Of Services.SyncDeletionProgress) = Nothing
            If syncDeletions Then
                syncProgress = New Progress(Of Services.SyncDeletionProgress)(
                    Sub(p)
                        lblStatusCount.Text = String.Format("削除同期中... {0}/{1} ({2})",
                            p.ScannedCount, p.TotalCount, p.FolderName)
                    End Sub)
            End If

            Dim ct As System.Threading.CancellationToken = _importCts.Token

            Dim staThread As New System.Threading.Thread(
                Sub()
                    Try
                        Using outlookSvc As Services.OutlookService = Services.OutlookService.Connect()
                            Dim threadingSvc As New Services.ThreadingService(repo)
                            Dim ruleRepo As New Data.AutoDeleteRuleRepository(dbManager)
                            Dim autoDeleteSvc As New Services.AutoDeleteService(ruleRepo, repo)
                            Dim importSvc As New Services.ImportService(outlookSvc, repo, threadingSvc, settings, dbManager, autoDeleteSvc)
                            Dim importResult As Services.ImportResult = importSvc.ImportFolders(targetFolders, maxCount, progress, ct, fullSync)

                            ' 削除同期
                            If syncDeletions Then
                                importResult.DeletedCount = importSvc.SyncDeletionsForFolders(targetFolders, syncProgress, ct)
                            End If

                            tcs.SetResult(importResult)
                        End Using
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)
            staThread.SetApartmentState(System.Threading.ApartmentState.STA)
            staThread.IsBackground = True
            staThread.Start()

            Dim result As Services.ImportResult = Await tcs.Task
            importStopwatch.Stop()
            _lastOutlookTotalCount = result.TotalOutlookCount

            Dim importLogMsg As String = String.Format(
                "取り込みが完了しました — 取り込み: {0}件, スキップ: {1}件, エラー: {2}件, 所要時間: {3:F1}秒",
                result.ImportedCount, result.SkippedCount, result.ErrorCount,
                importStopwatch.Elapsed.TotalSeconds)
            If result.ErrorSkipCount > 0 Then
                importLogMsg &= String.Format(", エラー除外: {0}件", result.ErrorSkipCount)
            End If
            Services.Logger.Info(importLogMsg)

            If result.DeletedCount > 0 Then
                Services.Logger.Info(String.Format("削除同期: {0}件", result.DeletedCount))
            End If

            If result.ErrorCount > 0 Then
                Services.Logger.Warn(String.Format("取り込みエラーが {0}件 発生しました", result.ErrorCount))
            End If

            Dim msg As String = String.Format(
                "取り込み完了{0}取り込み: {1}件 / スキップ: {2}件 / エラー: {3}件",
                vbCrLf, result.ImportedCount, result.SkippedCount, result.ErrorCount)
            If result.ErrorSkipCount > 0 Then
                msg &= String.Format(" / エラー除外: {0}件", result.ErrorSkipCount)
            End If
            If result.AutoDeletedCount > 0 Then
                msg &= String.Format(" / 自動削除: {0}件", result.AutoDeletedCount)
            End If
            If result.DeletedCount > 0 Then
                msg &= String.Format("{0}削除同期: {1}件", vbCrLf, result.DeletedCount)
            End If
            If result.ErrorCount > 0 Then
                ' エラー詳細をログファイルに出力
                Dim logPath As String = Services.ImportLogWriter.WriteErrorLog(result.Errors)

                ' エラーメッセージ別の集計サマリーをダイアログに表示
                Dim summary As List(Of KeyValuePair(Of String, Integer)) =
                    Services.ImportLogWriter.SummarizeErrors(result.Errors)
                msg &= vbCrLf & vbCrLf & "エラーサマリー:"
                For Each kvp As KeyValuePair(Of String, Integer) In summary
                    msg &= vbCrLf & "  ・" & kvp.Key & ": " & kvp.Value.ToString() & "件"
                Next

                If logPath IsNot Nothing Then
                    msg &= vbCrLf & vbCrLf & "詳細はログファイルを参照:" & vbCrLf & logPath
                End If
            End If
            ' エラー時と正常時で別々の設定に従う
            If (result.ErrorCount > 0 AndAlso _settings.ShowImportErrorDialog) OrElse
               (result.ErrorCount = 0 AndAlso _settings.ShowImportResult) Then
                MessageBox.Show(msg, "取り込み結果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ' トレイ常駐中はバルーン通知
            ShowImportBalloon(result)

            ' ゴミ箱の期限切れメールをパージ
            PurgeExpiredTrash()

            ' フォルダツリーとメール一覧を更新
            LoadFolderTree()
            Await LoadEmailsAsync(_currentFolder)

        Catch ex As OperationCanceledException
            Services.Logger.Info("取り込みがユーザーにより中断されました")
            lblStatusCount.Text = "取り込みを中断しました"
        Catch ex As AggregateException When ex.InnerException IsNot Nothing AndAlso
                                             TypeOf ex.InnerException Is OperationCanceledException
            Services.Logger.Info("取り込みがユーザーにより中断されました")
            lblStatusCount.Text = "取り込みを中断しました"
        Catch ex As Exception
            Services.Logger.Error("取り込み中に予期しないエラーが発生しました", ex)
            MessageBox.Show("取り込みエラー:" & vbCrLf & ex.Message, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isImporting = False
            If _importCts IsNot Nothing Then
                _importCts.Dispose()
                _importCts = Nothing
            End If
            btnImportNow.Enabled = True
            menuItemImportNow.Enabled = True
            menuItemImportCancel.Enabled = False
        End Try
        Await UpdateStatusBarAsync()
    End Function

    ' ════════════════════════════════════════════════════════════
    '  タスクトレイ常駐
    ' ════════════════════════════════════════════════════════════

    ''' <summary>NotifyIcon のアイコンとイベントを設定する。</summary>
    Private Sub SetupNotifyIcon()
        ' 実行ファイルからアイコンを取得（アイコン未設定の場合はデフォルトアイコン）
        Dim exePath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim appIcon As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath)
        If appIcon IsNot Nothing Then
            notifyIcon.Icon = appIcon
            Me.Icon = appIcon
        End If
        notifyIcon.Visible = True

        AddHandler notifyIcon.DoubleClick, AddressOf NotifyIcon_DoubleClick
        AddHandler trayMenuShow.Click, AddressOf TrayMenuShow_Click
        AddHandler trayMenuImportNow.Click, AddressOf TrayMenuImportNow_Click
        AddHandler trayMenuExit.Click, AddressOf TrayMenuExit_Click
    End Sub

    ''' <summary>フォームをタスクトレイに格納する。</summary>
    Private Sub MinimizeToSystemTray()
        Me.Hide()
        Me.ShowInTaskbar = False
        Services.Logger.Info("タスクトレイに格納しました")
    End Sub

    ''' <summary>タスクトレイからフォームを復帰する。</summary>
    Private Sub RestoreFromSystemTray()
        Me.Show()
        Me.ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        Me.Activate()
    End Sub

    ''' <summary>最小化時にタスクトレイへ格納する。</summary>
    Private Sub MainForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized AndAlso _settings.MinimizeToTray Then
            MinimizeToSystemTray()
        End If
    End Sub

    Private Sub NotifyIcon_DoubleClick(sender As Object, e As EventArgs)
        RestoreFromSystemTray()
    End Sub

    Private Sub TrayMenuShow_Click(sender As Object, e As EventArgs)
        RestoreFromSystemTray()
    End Sub

    Private Async Sub TrayMenuImportNow_Click(sender As Object, e As EventArgs)
        RestoreFromSystemTray()
        Await RunImportAsync()
    End Sub

    Private Sub TrayMenuExit_Click(sender As Object, e As EventArgs)
        _isRealClose = True
        Me.Close()
    End Sub

    ''' <summary>取り込み完了時にバルーン通知を表示する。</summary>
    Private Sub ShowImportBalloon(result As Services.ImportResult)
        If Not _settings.ShowBalloonOnImport Then Return
        If Not notifyIcon.Visible Then Return

        ' 取り込み0件・削除同期0件・エラー0件の場合は通知不要
        If result.ImportedCount = 0 AndAlso result.DeletedCount = 0 AndAlso result.ErrorCount = 0 Then Return

        Dim msg As String = String.Format("取り込み: {0}件 / スキップ: {1}件",
            result.ImportedCount, result.SkippedCount)
        If result.ErrorCount > 0 Then
            msg &= String.Format(" / エラー: {0}件", result.ErrorCount)
        End If
        If result.ErrorSkipCount > 0 Then
            msg &= String.Format(" / エラー除外: {0}件", result.ErrorSkipCount)
        End If
        If result.DeletedCount > 0 Then
            msg &= String.Format(" / 削除同期: {0}件", result.DeletedCount)
        End If

        Dim tipIcon As ToolTipIcon = If(result.ErrorCount > 0, ToolTipIcon.Warning, ToolTipIcon.Info)
        notifyIcon.ShowBalloonTip(5000, "取り込み完了", msg, tipIcon)
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  自動取り込み
    ' ════════════════════════════════════════════════════════════

    Private Sub StartAutoImport()
        _autoImportEnabled = True
        If _settings.AutoImportMode = 1 Then
            ' 定時モード
            _autoImportTimer.Stop()
            _scheduledImportTimer.Start()
            Dim timeStr As String = _settings.ScheduledImportTime
            btnAutoImport.Text = "定時: " & timeStr
            btnAutoImport.ToolTipText = String.Format("定時取り込み {0}（クリックで停止）", timeStr)
        Else
            ' 間隔モード
            _scheduledImportTimer.Stop()
            _autoImportTimer.Start()
            btnAutoImport.Text = "自動: ●"
            btnAutoImport.ToolTipText = "自動取り込み停止（クリックで停止）"
        End If
    End Sub

    Private Sub StopAutoImport()
        _autoImportEnabled = False
        _autoImportTimer.Stop()
        _scheduledImportTimer.Stop()
        btnAutoImport.Text = "自動: ▶"
        btnAutoImport.ToolTipText = "自動取り込み開始（クリックで開始）"
    End Sub

    Private Sub btnAutoImport_Click(sender As Object, e As EventArgs) Handles btnAutoImport.Click
        If _autoImportEnabled Then
            StopAutoImport()
        Else
            StartAutoImport()
        End If
    End Sub

    Private Async Sub AutoImportTimer_Tick(sender As Object, e As EventArgs) Handles _autoImportTimer.Tick
        Await RunImportAsync()
    End Sub

    ''' <summary>定時取り込み: 1分間隔で現在時刻と設定時刻を比較し、一致したら完全同期を実行する。</summary>
    Private Async Sub ScheduledImportTimer_Tick(sender As Object, e As EventArgs) Handles _scheduledImportTimer.Tick
        Dim nowTime As String = DateTime.Now.ToString("HH:mm")
        If nowTime <> _settings.ScheduledImportTime Then Return

        ' 同じ日に2回実行しない
        If _lastScheduledRunDate.Date = DateTime.Today Then Return
        _lastScheduledRunDate = DateTime.Today

        Services.Logger.Info(String.Format("定時取り込みを開始します（{0}）", nowTime))
        Await RunImportAsync(fullSync:=True)
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  検索
    ' ════════════════════════════════════════════════════════════

    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.F Then
            txtSearch.Focus()
            txtSearch.SelectAll()
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        RunSearch()
    End Sub


    Private Sub btnClearSearch_Click(sender As Object, e As EventArgs) Handles btnClearSearch.Click
        txtSearch.Text = String.Empty
        RunSearch()
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            RunSearch()
        End If
    End Sub

    Private Async Sub RunSearch()
        Dim query As String = txtSearch.Text.Trim()
        _searchQuery = If(String.IsNullOrEmpty(query), Nothing, query)

        If String.IsNullOrEmpty(query) Then
            ' クエリ空 → 現在のフォルダを再表示
            Await LoadEmailsAsync(_currentFolder)
            Return
        End If

        Try
            Dim repo As Data.EmailRepository = _repo
            _emailCache = Await Task.Run(
                Function() As List(Of Models.Email)
                    Return repo.SearchEmailsFiltered(query)
                End Function)
            SortEmailCache()
            listViewEmails.VirtualListSize = _emailCache.Count
            listViewEmails.Invalidate()
            ' プレビューをクリア（新しい検索結果を選択するまで）
            emailPreview.ClearPreview()
            conversationView.ClearView()
            lblStatusCount.Text = String.Format("検索結果: {0:N0}件", _emailCache.Count)
            UpdateFolderCountLabel()
        Catch ex As Exception
            MessageBox.Show("検索エラー:" & vbCrLf & ex.Message, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  フォルダ件数ラベル
    ' ════════════════════════════════════════════════════════════

    ''' <summary>ToolStrip 右端のフォルダ件数ラベルを更新する。</summary>
    Private Sub UpdateFolderCountLabel()
        Dim folderDisplayName As String = If(_currentFolder, "すべて")

        If _searchQuery IsNot Nothing Then
            ' 検索フィルタ中: "フォルダ名 N / M 件"
            lblFolderCount.Text = String.Format("{0} {1:N0} / {2:N0} 件",
                folderDisplayName, _emailCache.Count, _currentFolderTotalCount)
        Else
            ' 通常表示: "フォルダ名 N 件"
            lblFolderCount.Text = String.Format("{0} {1:N0} 件",
                folderDisplayName, _emailCache.Count)
        End If
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  削除
    ' ════════════════════════════════════════════════════════════

    ''' <summary>メール一覧で選択中のメールをゴミ箱に移動する（論理削除）。</summary>
    Private Async Sub DeleteSelectedEmails()
        If listViewEmails.SelectedIndices.Count = 0 Then Return

        Dim selectedIds As New List(Of Integer)()
        For Each idx As Integer In listViewEmails.SelectedIndices
            If idx >= 0 AndAlso idx < _emailCache.Count Then
                selectedIds.Add(_emailCache(idx).Id)
            End If
        Next
        If selectedIds.Count = 0 Then Return

        Dim msg As String
        If selectedIds.Count = 1 Then
            msg = "選択したメールをゴミ箱に移動しますか？"
        Else
            msg = String.Format("選択した {0} 件のメールをゴミ箱に移動しますか？", selectedIds.Count)
        End If

        If MessageBox.Show(msg, "ゴミ箱に移動",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question,
                           MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Return
        End If

        emailPreview.ClearPreview()
        conversationView.ClearView()

        _repo.SoftDeleteEmailsByIds(selectedIds)
        Await LoadEmailsAsync(_currentFolder)
        LoadFolderTree()
    End Sub

    ''' <summary>ゴミ箱で選択中のメールを復元する。</summary>
    Private Async Sub RestoreSelectedEmails()
        If listViewEmails.SelectedIndices.Count = 0 Then Return

        Dim selectedIds As New List(Of Integer)()
        For Each idx As Integer In listViewEmails.SelectedIndices
            If idx >= 0 AndAlso idx < _emailCache.Count Then
                selectedIds.Add(_emailCache(idx).Id)
            End If
        Next
        If selectedIds.Count = 0 Then Return

        _repo.RestoreEmailsByIds(selectedIds)

        emailPreview.ClearPreview()
        conversationView.ClearView()
        Await LoadEmailsAsync(_currentFolder)
        LoadFolderTree()
    End Sub

    ''' <summary>ゴミ箱で選択中のメールを完全に削除する（物理削除）。</summary>
    Private Async Sub PurgeSelectedEmails()
        If listViewEmails.SelectedIndices.Count = 0 Then Return

        Dim selectedIds As New List(Of Integer)()
        For Each idx As Integer In listViewEmails.SelectedIndices
            If idx >= 0 AndAlso idx < _emailCache.Count Then
                selectedIds.Add(_emailCache(idx).Id)
            End If
        Next
        If selectedIds.Count = 0 Then Return

        Dim msg As String
        If selectedIds.Count = 1 Then
            msg = "選択したメールを完全に削除しますか？" & vbCrLf &
                  "添付ファイルも削除されます。この操作は元に戻せません。"
        Else
            msg = String.Format("選択した {0} 件のメールを完全に削除しますか？" & vbCrLf &
                  "添付ファイルも削除されます。この操作は元に戻せません。", selectedIds.Count)
        End If

        If MessageBox.Show(msg, "完全削除の確認",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning,
                           MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Return
        End If

        emailPreview.ClearPreview()
        conversationView.ClearView()

        Dim attachmentPaths As List(Of String) = _repo.PurgeEmailsByIds(selectedIds)
        For Each filePath As String In attachmentPaths
            If IO.File.Exists(filePath) Then
                Try
                    IO.File.Delete(filePath)
                Catch
                End Try
            End If
        Next

        Await LoadEmailsAsync(_currentFolder)
        LoadFolderTree()
    End Sub

    ''' <summary>ゴミ箱を空にする。</summary>
    Private Async Sub EmptyTrash()
        Dim trashCount As Integer = _repo.GetTrashCount()
        If trashCount = 0 Then
            MessageBox.Show("ゴミ箱は空です。", "ゴミ箱を空にする", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If MessageBox.Show(
                String.Format("ゴミ箱内の {0} 件のメールを完全に削除しますか？" & vbCrLf &
                              "添付ファイルも削除されます。この操作は元に戻せません。", trashCount),
                "ゴミ箱を空にする",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Return
        End If

        emailPreview.ClearPreview()
        conversationView.ClearView()

        ' ゴミ箱内の全メール ID を取得して物理削除
        Dim trashEmails As List(Of Models.Email) = _repo.GetTrashEmails()
        Dim trashIds As New List(Of Integer)()
        For Each email As Models.Email In trashEmails
            trashIds.Add(email.Id)
        Next

        Dim attachmentPaths As List(Of String) = _repo.PurgeEmailsByIds(trashIds)
        For Each filePath As String In attachmentPaths
            If IO.File.Exists(filePath) Then
                Try
                    IO.File.Delete(filePath)
                Catch
                End Try
            End If
        Next

        Await LoadEmailsAsync(_currentFolder)
        LoadFolderTree()
    End Sub

    ''' <summary>起動時にゴミ箱の期限切れメールを自動パージする。</summary>
    Private Sub PurgeExpiredTrash()
        Dim days As Integer = _settings.TrashAutoDeleteDays
        If days <= 0 Then Return

        Try
            Dim attachmentPaths As List(Of String) = _repo.PurgeExpiredTrash(days)
            If attachmentPaths.Count > 0 Then
                For Each filePath As String In attachmentPaths
                    If IO.File.Exists(filePath) Then
                        Try
                            IO.File.Delete(filePath)
                        Catch
                        End Try
                    End If
                Next
                Services.Logger.Info(String.Format("ゴミ箱の自動パージが完了しました（{0}日経過分）", days))
            End If
        Catch ex As Exception
            Services.Logger.Warn("ゴミ箱の自動パージに失敗しました: " & ex.Message)
        End Try
    End Sub

    ''' <summary>コンテキストメニューをゴミ箱表示と通常表示で切り替える。</summary>
    Private Sub UpdateContextMenuForView()
        listViewContextMenu.Items.Clear()
        If _isTrashView Then
            listViewContextMenu.Items.Add(restoreMenuItem)
            listViewContextMenu.Items.Add(purgeMenuItem)
        Else
            listViewContextMenu.Items.Add(deleteMenuItem)
        End If
        UpdateTrashFolderColumn()
    End Sub

    ''' <summary>ゴミ箱表示時に追加列（フォルダ・削除日時・削除予定日時）を追加し、通常表示時に除去する。</summary>
    Private Sub UpdateTrashFolderColumn()
        Dim trashColNames() As String = {"colTrashFolder", "colDeletedAt", "colPurgeAt"}

        If _isTrashView Then
            ' 未追加の列のみ追加
            For Each colName As String In trashColNames
                Dim found As Boolean = False
                For Each col As ColumnHeader In listViewEmails.Columns
                    If col.Name = colName Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then
                    Dim col As New ColumnHeader()
                    col.Name = colName
                    Select Case colName
                        Case "colTrashFolder"
                            col.Text = "フォルダ"
                            col.Width = 140
                        Case "colDeletedAt"
                            col.Text = "削除日時"
                            col.Width = 120
                        Case "colPurgeAt"
                            col.Text = "削除予定日時"
                            col.Width = 120
                    End Select
                    listViewEmails.Columns.Add(col)
                End If
            Next
        Else
            ' ゴミ箱用列を除去（逆順で削除）
            For i As Integer = listViewEmails.Columns.Count - 1 To 0 Step -1
                If Array.IndexOf(trashColNames, listViewEmails.Columns(i).Name) >= 0 Then
                    listViewEmails.Columns.RemoveAt(i)
                End If
            Next
        End If
    End Sub

    ''' <summary>コンテキストメニュー「削除」クリック。</summary>
    Private Sub OnDeleteEmailMenuClick(sender As Object, e As EventArgs) Handles deleteMenuItem.Click
        DeleteSelectedEmails()
    End Sub

    ''' <summary>コンテキストメニュー「復元」クリック。</summary>
    Private Sub OnRestoreEmailMenuClick(sender As Object, e As EventArgs) Handles restoreMenuItem.Click
        RestoreSelectedEmails()
    End Sub

    ''' <summary>コンテキストメニュー「完全に削除」クリック。</summary>
    Private Sub OnPurgeEmailMenuClick(sender As Object, e As EventArgs) Handles purgeMenuItem.Click
        PurgeSelectedEmails()
    End Sub

    ''' <summary>コンテキストメニュー「ゴミ箱を空にする」クリック。</summary>
    Private Sub OnEmptyTrashMenuClick(sender As Object, e As EventArgs) Handles emptyTrashMenuItem.Click
        EmptyTrash()
    End Sub

    ''' <summary>メール一覧で Delete キーが押されたときに選択メールを削除する。</summary>
    Private Sub listViewEmails_KeyDown(sender As Object, e As KeyEventArgs) Handles listViewEmails.KeyDown
        If e.KeyCode = Keys.Delete Then
            If _isTrashView Then
                PurgeSelectedEmails()
            Else
                DeleteSelectedEmails()
            End If
            e.Handled = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
            ' VirtualMode では Ctrl+A が自動で動作しないため Win32 API で一括全選択
            Dim lvi As New LVITEM()
            lvi.mask = LVIF_STATE
            lvi.state = LVIS_SELECTED
            lvi.stateMask = LVIS_SELECTED
            ' iItem = -1 で全アイテムに適用
            SendMessage(listViewEmails.Handle, LVM_SETITEMSTATE, New IntPtr(-1), lvi)
            e.Handled = True
        End If
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  テキスト/HTML 切替ボタン
    ' ════════════════════════════════════════════════════════════

    ''' <summary>テキスト/HTML切替ボタンをタブコントロール右上に配置する。</summary>
    Private Sub SetupToggleViewButton()
        btnToggleView.BringToFront()
        PositionToggleViewButton()
    End Sub

    Private Sub PositionToggleViewButton()
        ' タブ行の右端に配置
        btnToggleView.Top = 1
        btnToggleView.Left = tabControl.Width - btnToggleView.Width - 4
    End Sub

    Private Sub OnTabControlSizeChanged(sender As Object, e As EventArgs) Handles tabControl.SizeChanged
        PositionToggleViewButton()
    End Sub

    Private Sub OnToggleViewClick(sender As Object, e As EventArgs) Handles btnToggleView.Click
        emailPreview.ToggleView()
        UpdateToggleViewButton()
    End Sub

    Private Sub UpdateToggleViewButton()
        btnToggleView.Enabled = emailPreview.CanToggleView
        If emailPreview.IsHtmlView Then
            btnToggleView.Text = "テキスト表示"
        Else
            btnToggleView.Text = "HTML 表示"
        End If
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  フォルダ件数
    ' ════════════════════════════════════════════════════════════

    ''' <summary>フォルダツリーの件数表示を更新する。選択中のフォルダは維持する。</summary>
    Private Async Function UpdateFolderCountsAsync() As Task
        Dim repo As Data.EmailRepository = _repo

        ' DBクエリをバックグラウンドスレッドで実行（1クエリで全フォルダ件数を取得）
        Dim folderData As List(Of Tuple(Of String, Integer)) = Await Task.Run(
            Function() As List(Of Tuple(Of String, Integer))
                Dim result As New List(Of Tuple(Of String, Integer))()
                Try
                    Dim counts As Tuple(Of Integer, Dictionary(Of String, Integer)) = repo.GetFolderCounts()
                    ' 「すべて」を先頭に配置
                    result.Add(Tuple.Create(CType(Nothing, String), counts.Item1))
                    For Each kvp As KeyValuePair(Of String, Integer) In counts.Item2
                        result.Add(Tuple.Create(kvp.Key, kvp.Value))
                    Next
                    ' ゴミ箱を末尾に配置
                    result.Add(Tuple.Create(TrashFolderTag, repo.GetTrashCount()))
                Catch ex As Exception
                    Services.Logger.Warn("フォルダ件数の取得に失敗しました: " & ex.Message)
                End Try
                Return result
            End Function)

        ' UIスレッドでツリーを更新
        If folderData.Count = 0 Then Return

        _updatingFolderCounts = True
        Dim selectedTag As String = If(_isTrashView, TrashFolderTag, _currentFolder)
        treeViewFolders.BeginUpdate()
        treeViewFolders.Nodes.Clear()

        For Each item As Tuple(Of String, Integer) In folderData
            Dim displayName As String
            If item.Item1 Is Nothing Then
                displayName = String.Format("すべて ({0:N0})", item.Item2)
            ElseIf String.Equals(item.Item1, TrashFolderTag, StringComparison.Ordinal) Then
                displayName = String.Format("ゴミ箱 ({0:N0})", item.Item2)
            Else
                displayName = String.Format("{0} ({1:N0})", item.Item1, item.Item2)
            End If
            Dim node As New TreeNode(displayName)
            node.Tag = item.Item1
            treeViewFolders.Nodes.Add(node)
        Next

        ' 選択状態を復元
        Dim found As Boolean = False
        If selectedTag IsNot Nothing Then
            For Each node As TreeNode In treeViewFolders.Nodes
                If node.Tag IsNot Nothing AndAlso CStr(node.Tag) = selectedTag Then
                    treeViewFolders.SelectedNode = node
                    found = True
                    Exit For
                End If
            Next
        End If
        If Not found Then
            treeViewFolders.SelectedNode = treeViewFolders.Nodes(0)
        End If

        treeViewFolders.EndUpdate()
        _updatingFolderCounts = False
    End Function

    ' ════════════════════════════════════════════════════════════
    '  ステータスバー
    ' ════════════════════════════════════════════════════════════

    Private Async Function UpdateStatusBarAsync() As Task
        Dim repo As Data.EmailRepository = _repo
        Dim lastOutlookTotal As Integer = _lastOutlookTotalCount

        ' DBクエリをバックグラウンドスレッドで実行
        Dim result As Tuple(Of Integer, DateTime?) = Await Task.Run(
            Function() As Tuple(Of Integer, DateTime?)
                Dim count As Integer = 0
                Dim lastImport As DateTime? = Nothing
                Try
                    count = repo.GetTotalCount()
                Catch ex As Exception
                    Services.Logger.Warn("ステータスバーの総数取得に失敗しました: " & ex.Message)
                End Try
                Try
                    lastImport = repo.GetLastImportDate()
                Catch ex As Exception
                    Services.Logger.Warn("ステータスバーの最終取り込み日時取得に失敗しました: " & ex.Message)
                End Try
                Return Tuple.Create(count, lastImport)
            End Function)

        ' UIスレッドで表示を更新
        If result.Item1 > 0 OrElse lastOutlookTotal <= 0 Then
            If lastOutlookTotal > 0 AndAlso result.Item1 < lastOutlookTotal Then
                lblStatusCount.Text = String.Format("総数 {0:N0}/{1:N0}件", result.Item1, lastOutlookTotal)
            Else
                lblStatusCount.Text = String.Format("総数 {0:N0}件", result.Item1)
            End If
        Else
            lblStatusCount.Text = "総数 -件"
        End If

        If result.Item2.HasValue Then
            lblStatusLastImport.Text = "最終取り込み: " & result.Item2.Value.ToString("yyyy/MM/dd HH:mm")
        Else
            lblStatusLastImport.Text = "最終取り込み: -"
        End If

        Await UpdateFolderCountsAsync()
    End Function

    ' ════════════════════════════════════════════════════════════
    '  メール一覧 ソート・列設定
    ' ════════════════════════════════════════════════════════════

    ''' <summary>列名クリックでソート方向を切り替え、一覧を更新する。</summary>
    Private Sub listViewEmails_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles listViewEmails.ColumnClick
        Dim clickedColumn As EmailListColumn = CType(e.Column, EmailListColumn)
        If clickedColumn = _sortColumn Then
            _sortAscending = Not _sortAscending
        Else
            _sortColumn = clickedColumn
            ' 受信日時はデフォルト降順、他は昇順
            _sortAscending = (clickedColumn <> EmailListColumn.ReceivedAt)
        End If
        UpdateColumnSortIndicator()

        ' ソート前に選択中メールのIDを保存
        Dim selectedIds As New HashSet(Of Integer)()
        For Each idx As Integer In listViewEmails.SelectedIndices
            If idx >= 0 AndAlso idx < _emailCache.Count Then
                selectedIds.Add(_emailCache(idx).Id)
            End If
        Next

        SortEmailCache()

        ' ソート後: 選択をクリアし、保存したIDの新位置に再設定
        If selectedIds.Count > 0 Then
            ' Win32 API で全選択解除（VirtualMode で高速）
            Dim lviClear As LVITEM
            lviClear.state = 0
            lviClear.stateMask = LVIS_SELECTED Or LVIS_FOCUSED
            SendMessage(listViewEmails.Handle, LVM_SETITEMSTATE, New IntPtr(-1), lviClear)

            Dim firstNewIdx As Integer = -1
            For i As Integer = 0 To _emailCache.Count - 1
                If selectedIds.Contains(_emailCache(i).Id) Then
                    ' Win32 API で個別選択
                    Dim lviSet As LVITEM
                    lviSet.state = LVIS_SELECTED
                    lviSet.stateMask = LVIS_SELECTED
                    SendMessage(listViewEmails.Handle, LVM_SETITEMSTATE, New IntPtr(i), lviSet)
                    If firstNewIdx = -1 Then firstNewIdx = i
                End If
            Next

            ' 最初の選択アイテムにフォーカスを設定しスクロール
            If firstNewIdx >= 0 Then
                Dim lviFocus As LVITEM
                lviFocus.state = LVIS_FOCUSED
                lviFocus.stateMask = LVIS_FOCUSED
                SendMessage(listViewEmails.Handle, LVM_SETITEMSTATE, New IntPtr(firstNewIdx), lviFocus)
                listViewEmails.EnsureVisible(firstNewIdx)
            End If
        End If

        listViewEmails.Invalidate()
    End Sub

    ''' <summary>_emailCache を現在のソート設定で並び替える。</summary>
    Private Sub SortEmailCache()
        If _emailCache Is Nothing OrElse _emailCache.Count = 0 Then Return
        Dim asc As Boolean = _sortAscending
        Select Case _sortColumn
            Case EmailListColumn.Attachment
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = a.HasAttachments.CompareTo(b.HasAttachments)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case EmailListColumn.Subject
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = String.Compare(
                                         If(a.Subject, String.Empty),
                                         If(b.Subject, String.Empty),
                                         StringComparison.OrdinalIgnoreCase)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case EmailListColumn.Sender
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim sa As String = If(Not String.IsNullOrEmpty(a.SenderName),
                                                           a.SenderName,
                                                           If(a.SenderEmail, String.Empty))
                                     Dim sb As String = If(Not String.IsNullOrEmpty(b.SenderName),
                                                           b.SenderName,
                                                           If(b.SenderEmail, String.Empty))
                                     Dim cmp As Integer = String.Compare(sa, sb, StringComparison.OrdinalIgnoreCase)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case EmailListColumn.ReceivedAt
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = a.ReceivedAt.CompareTo(b.ReceivedAt)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case EmailListColumn.Size
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = a.EmailSize.CompareTo(b.EmailSize)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
        End Select
    End Sub

    ''' <summary>現在のソート列と方向をヘッダーテキストに反映する。</summary>
    Private Sub UpdateColumnSortIndicator()
        If listViewEmails.Columns.Count < _colBaseNames.Length Then Return
        For i As Integer = 0 To _colBaseNames.Length - 1
            Dim indicator As String = If(i = CInt(_sortColumn),
                                         If(_sortAscending, " ↑", " ↓"),
                                         String.Empty)
            listViewEmails.Columns(i).Text = _colBaseNames(i) & indicator
        Next
    End Sub

    ''' <summary>AppSettings から列幅を復元する。</summary>
    Private Sub RestoreColumnWidths()
        RestoreColumnWidthsFrom(_settings.EmailListColumnWidths)
    End Sub

    ''' <summary>カンマ区切りの列幅文字列から列幅を復元する。</summary>
    Private Sub RestoreColumnWidthsFrom(colWidths As String)
        If String.IsNullOrEmpty(colWidths) Then Return

        Dim parts() As String = colWidths.Split(","c)
        If parts.Length <> listViewEmails.Columns.Count Then Return

        For i As Integer = 0 To parts.Length - 1
            Dim w As Integer
            If Integer.TryParse(parts(i), w) AndAlso w > 0 Then
                listViewEmails.Columns(i).Width = w
            End If
        Next
    End Sub

    ''' <summary>AppSettings から列の表示順（DisplayIndex）を復元する。</summary>
    Private Sub RestoreColumnOrder()
        RestoreColumnOrderFrom(_settings.EmailListColumnOrder)
    End Sub

    ''' <summary>カンマ区切りの DisplayIndex 文字列から列順を復元する。</summary>
    Private Sub RestoreColumnOrderFrom(colOrder As String)
        If String.IsNullOrEmpty(colOrder) Then Return

        Dim parts() As String = colOrder.Split(","c)
        If parts.Length <> listViewEmails.Columns.Count Then Return

        Dim displayIndices(parts.Length - 1) As Integer
        For i As Integer = 0 To parts.Length - 1
            If Not Integer.TryParse(parts(i), displayIndices(i)) Then Return
        Next

        ' DisplayIndex は小さい順にセットしないと競合するため、ターゲット順に並べてから適用する
        Dim pairs As New List(Of Tuple(Of Integer, Integer))()
        For i As Integer = 0 To displayIndices.Length - 1
            pairs.Add(Tuple.Create(i, displayIndices(i)))
        Next
        pairs.Sort(Function(a As Tuple(Of Integer, Integer), b As Tuple(Of Integer, Integer)) As Integer
                       Return a.Item2.CompareTo(b.Item2)
                   End Function)
        For Each pair As Tuple(Of Integer, Integer) In pairs
            listViewEmails.Columns(pair.Item1).DisplayIndex = pair.Item2
        Next
    End Sub

    ''' <summary>列の表示順・列幅・ソート設定を AppSettings に保存する。</summary>
    Private Sub SaveColumnSettings()
        Dim count As Integer = listViewEmails.Columns.Count
        Dim displayIndices(count - 1) As String
        Dim widths(count - 1) As String
        For i As Integer = 0 To count - 1
            displayIndices(i) = listViewEmails.Columns(i).DisplayIndex.ToString()
            widths(i) = listViewEmails.Columns(i).Width.ToString()
        Next
        Dim orderStr As String = String.Join(",", displayIndices)
        Dim widthStr As String = String.Join(",", widths)

        ' グローバル設定に保存（後方互換）
        _settings.EmailListColumnOrder = orderStr
        _settings.EmailListColumnWidths = widthStr
        _settings.EmailListSortColumn = CInt(_sortColumn)
        _settings.EmailListSortAscending = _sortAscending

        ' フォルダ別設定にも保存
        SaveColumnSettingsForFolder(GetCurrentFolderKey())
    End Sub

    ''' <summary>現在のフォルダを識別するキーを返す。</summary>
    Private Function GetCurrentFolderKey() As String
        If _isTrashView Then Return "__TRASH__"
        Return If(_currentFolder, "__ALL__")
    End Function

    ''' <summary>現在の列設定を指定フォルダ用に保存する。</summary>
    Private Sub SaveColumnSettingsForFolder(folderKey As String)
        Dim count As Integer = listViewEmails.Columns.Count
        Dim displayIndices(count - 1) As String
        Dim widths(count - 1) As String
        For i As Integer = 0 To count - 1
            displayIndices(i) = listViewEmails.Columns(i).DisplayIndex.ToString()
            widths(i) = listViewEmails.Columns(i).Width.ToString()
        Next
        _settings.SaveFolderColumnSettings(
            folderKey,
            String.Join(",", widths),
            String.Join(",", displayIndices),
            CInt(_sortColumn),
            _sortAscending)
    End Sub

    ''' <summary>指定フォルダの列設定を復元する。フォルダ別設定がなければグローバル設定を使用する。</summary>
    Private Sub RestoreColumnSettingsForFolder(folderKey As String)
        If _settings.HasFolderColumnSettings(folderKey) Then
            RestoreColumnWidthsFrom(_settings.LoadFolderColumnWidths(folderKey))
            RestoreColumnOrderFrom(_settings.LoadFolderColumnOrder(folderKey))
            Dim sortCol As Integer = _settings.LoadFolderSortColumn(folderKey)
            If sortCol >= 0 AndAlso sortCol <= CInt(EmailListColumn.Size) Then
                _sortColumn = CType(sortCol, EmailListColumn)
                _sortAscending = _settings.LoadFolderSortAscending(folderKey)
            End If
        Else
            RestoreColumnWidths()
            RestoreColumnOrder()
        End If
        UpdateColumnSortIndicator()
    End Sub

    ''' <summary>フォーム終了時に列設定を保存する。取り込み中の場合は確認ダイアログを表示する。</summary>
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' 閉じるボタンでトレイに格納（ユーザー操作かつ本当の終了でない場合）
        If e.CloseReason = CloseReason.UserClosing AndAlso
           Not _isRealClose AndAlso
           _settings.CloseToTray Then
            e.Cancel = True
            MinimizeToSystemTray()
            Return
        End If

        If _isImporting Then
            Dim answer As DialogResult = MessageBox.Show(
                "メールの取り込み中です。中断して終了しますか？" & vbCrLf &
                "（処理中のバッチはロールバックされます）",
                "終了の確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2)

            If answer = DialogResult.No Then
                e.Cancel = True
                Return
            End If

            ' キャンセルを要求して取り込みスレッドの終了を待機
            If _importCts IsNot Nothing Then
                _importCts.Cancel()
            End If

            ' 取り込みスレッドが終了するまで最大5秒間待機
            Dim waitStart As DateTime = DateTime.Now
            Do While _isImporting AndAlso (DateTime.Now - waitStart).TotalMilliseconds < 5000
                System.Threading.Thread.Sleep(100)
                Application.DoEvents()
            Loop
        End If

        SaveColumnSettings()
        notifyIcon.Visible = False

        ' EmailRepository のバルクリソースを確実に解放
        If _repo IsNot Nothing Then
            _repo.Dispose()
        End If

        Services.Logger.Info("アプリケーションを終了します")
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  メニューハンドラ
    ' ════════════════════════════════════════════════════════════

    Private Sub menuItemFileExit_Click(sender As Object, e As EventArgs) Handles menuItemFileExit.Click
        _isRealClose = True
        Me.Close()
    End Sub

    Private Sub menuItemSettings_Click(sender As Object, e As EventArgs) Handles menuItemSettings.Click
        OpenSettingsDialog()
    End Sub

    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click
        OpenSettingsDialog()
    End Sub

    ''' <summary>設定ダイアログを表示し、結果に応じてアプリケーション設定を反映する。</summary>
    Private Sub OpenSettingsDialog()
        Dim wasReset As Boolean = False
        Using frm As New SettingsForm()
            frm.ShowDialog(Me)
            wasReset = frm.DataWasReset
        End Using

        If wasReset Then
            ReinitializeApp()
        Else
            ' OK/キャンセルどちらでも設定を再適用（設定が変わった場合のみ影響あり）
            ApplySettings()
        End If
    End Sub

    ''' <summary>
    ''' データ初期化後にサービスと UI を再初期化する。
    ''' DB と添付ファイルはすでに削除済みの前提で呼び出すこと。
    ''' </summary>
    Private Async Sub ReinitializeApp()
        ' 既存リソースを解放してからサービス再生成
        If _repo IsNot Nothing Then
            _repo.Dispose()
        End If
        InitializeServices()

        ' 自動取り込みタイマー間隔を再設定
        ApplySettings()

        ' UI をリセット
        emailPreview.ClearPreview()
        conversationView.ClearView()
        _emailCache = New List(Of Models.Email)()
        _currentFolder = Nothing
        _searchQuery = Nothing
        txtSearch.Text = String.Empty

        LoadFolderTree()
        Await LoadEmailsAsync(Nothing)

        MessageBox.Show("データを初期化しました。",
            "初期化完了", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>設定変更をランタイムに反映する（タイマー間隔・自動取り込み開始/停止）。</summary>
    Private Sub ApplySettings()
        _autoImportTimer.Interval = _settings.AutoImportIntervalMinutes * 60 * 1000
        If _settings.AutoImportEnabled AndAlso Not _autoImportEnabled Then
            StartAutoImport()
        ElseIf Not _settings.AutoImportEnabled AndAlso _autoImportEnabled Then
            StopAutoImport()
        ElseIf _settings.AutoImportEnabled AndAlso _autoImportEnabled Then
            ' モードや時刻が変わった可能性があるため再起動
            StopAutoImport()
            StartAutoImport()
        End If
    End Sub

    ' ── エラー除外リスト ──────────────────────────────────────

    Private Sub menuItemErrorExclusion_Click(sender As Object, e As EventArgs) Handles menuItemErrorExclusion.Click
        Using frm As New Forms.ErrorExclusionForm(_repo)
            frm.ShowDialog(Me)
        End Using
    End Sub

    ' ── 自動削除ルール ──────────────────────────────────────

    Private Sub menuItemAutoDeleteRule_Click(sender As Object, e As EventArgs) Handles menuItemAutoDeleteRule.Click
        Dim ruleRepo As New Data.AutoDeleteRuleRepository(_dbManager)
        Dim autoDeleteSvc As New Services.AutoDeleteService(ruleRepo, _repo)
        Using frm As New Forms.AutoDeleteRuleForm(ruleRepo, _repo, autoDeleteSvc)
            frm.ShowDialog(Me)
        End Using
        ' ルール再適用でメールがゴミ箱に移動した可能性があるため更新
        LoadFolderTree()
    End Sub

    ' ── テーブルビューア ──────────────────────────────────────

    Private Sub menuItemDevTableViewer_Click(sender As Object, e As EventArgs) Handles menuItemDevTableViewer.Click
        ShowTableViewer("emails")
    End Sub

    Private Sub ShowTableViewer(tableName As String)
        Using frm As New Forms.TableViewerForm(_dbManager, tableName)
            frm.ShowDialog(Me)
        End Using
    End Sub

    ' ── 添付ファイル拡張子統計 ────────────────────────────────

    Private Sub menuItemDevAttachmentStats_Click(sender As Object, e As EventArgs) Handles menuItemDevAttachmentStats.Click
        Using frm As New Forms.AttachmentStatsForm(_dbManager)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub menuItemHelpManual_Click(sender As Object, e As EventArgs) Handles menuItemHelpManual.Click
        Dim frm As New Forms.HelpForm()
        frm.Show(Me)
    End Sub

    Private Sub menuItemHelpAbout_Click(sender As Object, e As EventArgs) Handles menuItemHelpAbout.Click
        Using frm As New Forms.AboutForm()
            frm.ShowDialog(Me)
        End Using
    End Sub

End Class
