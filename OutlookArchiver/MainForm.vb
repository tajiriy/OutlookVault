Option Explicit On
Option Strict On
Option Infer Off

Imports System.Threading.Tasks

Public Class MainForm

    ' ════════════════════════════════════════════════════════════
    '  フィールド
    ' ════════════════════════════════════════════════════════════

    Private _settings As Config.AppSettings
    Private _dbManager As Data.DatabaseManager
    Private _repo As Data.EmailRepository
    Private _autoImportTimer As System.Windows.Forms.Timer

    Private _emailCache As List(Of Models.Email)
    Private _currentFolder As String      ' Nothing = すべて
    Private _isImporting As Boolean
    Private _lastOutlookTotalCount As Integer = -1  ' -1 = 未取得
    Private _autoImportEnabled As Boolean
    Private _searchQuery As String        ' 現在の検索クエリ（Nothing = 検索なし）
    Private _updatingFolderCounts As Boolean ' フォルダ件数更新中のイベント抑制フラグ

    ' 列インデックス: 0=添付 1=件名 2=差出人 3=受信日時 4=サイズ
    Private _sortColumn As Integer = 3          ' デフォルト: 受信日時
    Private _sortAscending As Boolean = False   ' デフォルト: 降順
    Private ReadOnly _colBaseNames As String() = {"添付", "件名", "差出人", "受信日時", "サイズ"}

    ' ════════════════════════════════════════════════════════════
    '  フォーム初期化
    ' ════════════════════════════════════════════════════════════

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' デザイナーが SplitterDistance を上書きするため、ここで正しい比率に設定する
        ' フォルダツリー幅: 150px、メール一覧高さ: コンテンツ領域の約 40%
        splitMain.SplitterDistance = 150
        Dim contentHeight As Integer = Me.ClientSize.Height - menuStrip.Height - toolStrip.Height - statusStrip.Height
        splitRight.SplitterDistance = CInt(contentHeight * 0.4)

        InitializeServices()
        SetupAutoImportTimer()
        SetupEmailListColumns()
        SetupListViewContextMenu()
        SetupToggleViewButton()
        LoadFolderTree()
        UpdateStatusBar()

        If _settings.AutoImportEnabled Then
            StartAutoImport()
        End If
    End Sub

    Private Sub InitializeServices()
        _settings = Config.AppSettings.Instance
        _dbManager = New Data.DatabaseManager(_settings.DbFilePath)
        _dbManager.Initialize()
        _repo = New Data.EmailRepository(_dbManager)
        _emailCache = New List(Of Models.Email)()
    End Sub

    ''' <summary>メール一覧の列を設定し、列順・ソート設定を復元する。</summary>
    Private Sub SetupEmailListColumns()
        ' ── 添付列（index 0 に挿入 → SmallImageList アイコンが添付列に表示される）──
        Dim colAttach As New ColumnHeader()
        colAttach.Text = "添付"
        colAttach.Width = 30
        listViewEmails.Columns.Insert(0, colAttach)

        ' ── サイズ列（末尾に追加、index 4）─────────────────────────
        Dim colSize As New ColumnHeader()
        colSize.Text = "サイズ"
        colSize.Width = 80
        colSize.TextAlign = HorizontalAlignment.Right
        listViewEmails.Columns.Add(colSize)

        ' ── SmallImageList（添付アイコン）──────────────────────────
        Dim imgList As New ImageList()
        imgList.ImageSize = New Drawing.Size(16, 16)
        imgList.ColorDepth = ColorDepth.Depth32Bit
        imgList.Images.Add(New Drawing.Bitmap(16, 16))  ' index 0: 空白（添付なし）
        imgList.Images.Add(CreatePaperclipIcon())       ' index 1: ペーパークリップ
        listViewEmails.SmallImageList = imgList

        ' ── 列の並び替えを許可 ────────────────────────────────────
        listViewEmails.AllowColumnReorder = True

        ' ── ソート設定を復元 ──────────────────────────────────────
        _sortColumn = _settings.EmailListSortColumn
        _sortAscending = _settings.EmailListSortAscending

        ' ── 列幅を復元 ────────────────────────────────────────────
        RestoreColumnWidths()

        ' ── 列順を復元 ────────────────────────────────────────────
        RestoreColumnOrder()

        ' ── ソートインジケーターを更新 ───────────────────────────
        UpdateColumnSortIndicator()
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
        If sizeBytes >= 1024L * 1024L Then
            Return String.Format("{0:F1} MB", sizeBytes / (1024.0 * 1024.0))
        End If
        Dim kb As Long = (sizeBytes + 1023L) \ 1024L
        Return String.Format("{0} KB", Math.Max(1L, kb))
    End Function

    Private Sub SetupAutoImportTimer()
        _autoImportTimer = New System.Windows.Forms.Timer()
        _autoImportTimer.Interval = _settings.AutoImportIntervalMinutes * 60 * 1000
        AddHandler _autoImportTimer.Tick, AddressOf AutoImportTimer_Tick
    End Sub

    ''' <summary>メール一覧の複数選択・コンテキストメニューをコードで設定する。</summary>
    Private Sub SetupListViewContextMenu()
        listViewEmails.MultiSelect = True

        Dim ctxMenu As New System.Windows.Forms.ContextMenuStrip()
        Dim deleteItem As New System.Windows.Forms.ToolStripMenuItem("削除(&D)")
        AddHandler deleteItem.Click, AddressOf OnDeleteEmailMenuClick
        ctxMenu.Items.Add(CType(deleteItem, System.Windows.Forms.ToolStripItem))
        listViewEmails.ContextMenuStrip = ctxMenu
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  フォルダツリー
    ' ════════════════════════════════════════════════════════════

    Private Sub LoadFolderTree()
        treeViewFolders.BeginUpdate()
        treeViewFolders.Nodes.Clear()

        Dim totalCount As Integer = _repo.GetTotalCount()
        Dim nodeAll As New TreeNode(String.Format("すべて ({0})", totalCount))
        nodeAll.Tag = Nothing
        treeViewFolders.Nodes.Add(nodeAll)

        Dim folders As List(Of String) = _repo.GetFolderNames()
        For Each folder As String In folders
            Dim count As Integer = _repo.GetTotalCount(folder)
            Dim node As New TreeNode(String.Format("{0} ({1})", folder, count))
            node.Tag = folder
            treeViewFolders.Nodes.Add(node)
        Next

        treeViewFolders.EndUpdate()
        treeViewFolders.SelectedNode = nodeAll
    End Sub

    Private Sub treeViewFolders_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treeViewFolders.AfterSelect
        If e.Node Is Nothing Then Return
        If _updatingFolderCounts Then Return
        _currentFolder = TryCast(e.Node.Tag, String)
        ' 検索ボックスと検索クエリをクリアしてフォルダ切り替え
        txtSearch.Text = String.Empty
        _searchQuery = Nothing
        LoadEmails(_currentFolder)
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  メール一覧（VirtualMode）
    ' ════════════════════════════════════════════════════════════

    ''' <summary>DBからメールを読み込んでListViewのVirtualListSizeを更新する。</summary>
    Private Sub LoadEmails(Optional folderName As String = Nothing)
        _emailCache = _repo.GetEmails(folderName, pageSize:=500)
        SortEmailCache()
        listViewEmails.VirtualListSize = _emailCache.Count
        listViewEmails.Invalidate()
        UpdateStatusBar()
    End Sub

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
        item.Tag = email.Id
        e.Item = item
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  取り込み
    ' ════════════════════════════════════════════════════════════

    Private Async Sub btnImportNow_Click(sender As Object, e As EventArgs) Handles btnImportNow.Click, menuItemImportNow.Click
        Await RunImportAsync()
    End Sub

    ''' <summary>
    ''' Outlook COM は STA スレッドが必要なため、TaskCompletionSource + 専用STAスレッドで実行する。
    ''' </summary>
    Private Async Function RunImportAsync() As Task
        If _isImporting Then Return
        _isImporting = True
        btnImportNow.Enabled = False
        menuItemImportNow.Enabled = False
        lblStatusCount.Text = "取り込み中..."

        Try
            Dim progress As New Progress(Of Services.ImportProgress)(
                Sub(p)
                    lblStatusCount.Text = String.Format("取り込み中... {0}/{1} ({2})",
                        p.ProcessedCount, p.TotalCount, p.FolderName)
                End Sub)

            Dim tcs As New TaskCompletionSource(Of Services.ImportResult)()
            Dim targetFolders As List(Of String) = _settings.TargetFolders
            Dim maxCount As Integer = _settings.MaxImportCount
            Dim repo As Data.EmailRepository = _repo
            Dim settings As Config.AppSettings = _settings
            Dim dbManager As Data.DatabaseManager = _dbManager

            Dim staThread As New System.Threading.Thread(
                Sub()
                    Try
                        Using outlookSvc As Services.OutlookService = Services.OutlookService.Connect()
                            Dim threadingSvc As New Services.ThreadingService(repo)
                            Dim importSvc As New Services.ImportService(outlookSvc, repo, threadingSvc, settings, dbManager)
                            tcs.SetResult(importSvc.ImportFolders(targetFolders, maxCount, progress))
                        End Using
                    Catch ex As Exception
                        tcs.SetException(ex)
                    End Try
                End Sub)
            staThread.SetApartmentState(System.Threading.ApartmentState.STA)
            staThread.IsBackground = True
            staThread.Start()

            Dim result As Services.ImportResult = Await tcs.Task
            _lastOutlookTotalCount = result.TotalOutlookCount

            Dim msg As String = String.Format(
                "取り込み完了{0}取り込み: {1}件 / スキップ: {2}件 / エラー: {3}件",
                vbCrLf, result.ImportedCount, result.SkippedCount, result.ErrorCount)
            If result.ErrorCount > 0 Then
                msg &= vbCrLf & vbCrLf & "エラー詳細:" & vbCrLf & String.Join(vbCrLf, result.Errors)
            End If
            ' エラーがある場合は設定に関わらず表示。それ以外は ShowImportResult 設定に従う
            If result.ErrorCount > 0 OrElse _settings.ShowImportResult Then
                MessageBox.Show(msg, "取り込み結果", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ' フォルダツリーとメール一覧を更新
            LoadFolderTree()
            LoadEmails(_currentFolder)

        Catch ex As Exception
            MessageBox.Show("取り込みエラー:" & vbCrLf & ex.Message, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isImporting = False
            btnImportNow.Enabled = True
            menuItemImportNow.Enabled = True
            UpdateStatusBar()
        End Try
    End Function

    ' ════════════════════════════════════════════════════════════
    '  自動取り込み
    ' ════════════════════════════════════════════════════════════

    Private Sub StartAutoImport()
        _autoImportEnabled = True
        _autoImportTimer.Start()
        btnAutoImport.Text = "自動: ●"
        btnAutoImport.ToolTipText = "自動取り込み停止（クリックで停止）"
    End Sub

    Private Sub StopAutoImport()
        _autoImportEnabled = False
        _autoImportTimer.Stop()
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

    Private Async Sub AutoImportTimer_Tick(sender As Object, e As EventArgs)
        Await RunImportAsync()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  検索
    ' ════════════════════════════════════════════════════════════

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        RunSearch()
    End Sub

    Private Sub menuItemSearch_Click(sender As Object, e As EventArgs) Handles menuItemSearch.Click
        ' 検索ボックスにフォーカスを移してユーザーが入力できる状態にする
        txtSearch.Focus()
        txtSearch.SelectAll()
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            RunSearch()
        End If
    End Sub

    Private Sub RunSearch()
        Dim query As String = txtSearch.Text.Trim()
        _searchQuery = If(String.IsNullOrEmpty(query), Nothing, query)

        If String.IsNullOrEmpty(query) Then
            ' クエリ空 → 現在のフォルダを再表示
            LoadEmails(_currentFolder)
            Return
        End If

        Try
            _emailCache = _repo.SearchEmails(query)
            SortEmailCache()
            listViewEmails.VirtualListSize = _emailCache.Count
            listViewEmails.Invalidate()
            ' プレビューをクリア（新しい検索結果を選択するまで）
            emailPreview.ClearPreview()
            conversationView.ClearView()
            lblStatusCount.Text = String.Format("検索結果: {0}件", _emailCache.Count)
        Catch ex As Exception
            MessageBox.Show("検索エラー:" & vbCrLf & ex.Message, "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  削除
    ' ════════════════════════════════════════════════════════════

    ''' <summary>メール一覧で選択中のメールを削除する（添付ファイル実体も削除）。</summary>
    Private Sub DeleteSelectedEmails()
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
            msg = "選択したメールを削除しますか？" & vbCrLf &
                  "添付ファイルも削除されます。この操作は元に戻せません。"
        Else
            msg = String.Format("選択した {0} 件のメールを削除しますか？" & vbCrLf &
                  "添付ファイルも削除されます。この操作は元に戻せません。", selectedIds.Count)
        End If

        If MessageBox.Show(msg, "削除の確認",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning,
                           MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
            Return
        End If

        ' プレビューをクリア
        emailPreview.ClearPreview()
        conversationView.ClearView()

        For Each emailId As Integer In selectedIds
            ' 添付ファイルの物理削除
            Dim attachments As List(Of Models.Attachment) = _repo.GetAttachmentsByEmailId(emailId)
            For Each att As Models.Attachment In attachments
                If IO.File.Exists(att.FilePath) Then
                    Try
                        IO.File.Delete(att.FilePath)
                    Catch
                        ' 物理削除失敗は無視して DB 削除を継続
                    End Try
                End If
            Next
            ' DB から削除（トゥームストーン記録を含む）
            _repo.DeleteEmail(emailId)
        Next

        LoadEmails(_currentFolder)
        UpdateStatusBar()
    End Sub

    ''' <summary>コンテキストメニュー「削除」クリック。</summary>
    Private Sub OnDeleteEmailMenuClick(sender As Object, e As EventArgs)
        DeleteSelectedEmails()
    End Sub

    ''' <summary>メール一覧で Delete キーが押されたときに選択メールを削除する。</summary>
    Private Sub listViewEmails_KeyDown(sender As Object, e As KeyEventArgs) Handles listViewEmails.KeyDown
        If e.KeyCode = Keys.Delete Then
            DeleteSelectedEmails()
            e.Handled = True
        End If
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  テキスト/HTML 切替ボタン
    ' ════════════════════════════════════════════════════════════

    ''' <summary>テキスト/HTML切替ボタンをタブコントロール右上に配置する。</summary>
    Private Sub SetupToggleViewButton()
        btnToggleView.BringToFront()
        AddHandler btnToggleView.Click, AddressOf OnToggleViewClick
        AddHandler tabControl.SizeChanged, AddressOf OnTabControlSizeChanged
        PositionToggleViewButton()
    End Sub

    Private Sub PositionToggleViewButton()
        ' タブ行の右端に配置
        btnToggleView.Top = 1
        btnToggleView.Left = tabControl.Width - btnToggleView.Width - 4
    End Sub

    Private Sub OnTabControlSizeChanged(sender As Object, e As EventArgs)
        PositionToggleViewButton()
    End Sub

    Private Sub OnToggleViewClick(sender As Object, e As EventArgs)
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
    Private Sub UpdateFolderCounts()
        _updatingFolderCounts = True
        Dim selectedTag As String = _currentFolder
        treeViewFolders.BeginUpdate()
        treeViewFolders.Nodes.Clear()

        Try
            Dim totalCount As Integer = _repo.GetTotalCount()
            Dim nodeAll As New TreeNode(String.Format("すべて ({0})", totalCount))
            nodeAll.Tag = Nothing
            treeViewFolders.Nodes.Add(nodeAll)

            Dim folders As List(Of String) = _repo.GetFolderNames()
            For Each folder As String In folders
                Dim count As Integer = _repo.GetTotalCount(folder)
                Dim node As New TreeNode(String.Format("{0} ({1})", folder, count))
                node.Tag = folder
                treeViewFolders.Nodes.Add(node)
            Next

            ' 選択状態を復元
            Dim found As Boolean = False
            If selectedTag IsNot Nothing Then
                For Each node As TreeNode In treeViewFolders.Nodes
                    If CStr(node.Tag) = selectedTag Then
                        treeViewFolders.SelectedNode = node
                        found = True
                        Exit For
                    End If
                Next
            End If
            If Not found Then
                treeViewFolders.SelectedNode = treeViewFolders.Nodes(0)
            End If
        Catch
        End Try
        treeViewFolders.EndUpdate()
        _updatingFolderCounts = False
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  ステータスバー
    ' ════════════════════════════════════════════════════════════

    Private Sub UpdateStatusBar()
        Try
            Dim count As Integer = _repo.GetTotalCount()
            If _lastOutlookTotalCount > 0 AndAlso count < _lastOutlookTotalCount Then
                lblStatusCount.Text = String.Format("総数 {0}/{1}件", count, _lastOutlookTotalCount)
            Else
                lblStatusCount.Text = String.Format("総数 {0}件", count)
            End If
        Catch
            lblStatusCount.Text = "総数 -件"
        End Try
        Try
            Dim lastImport As DateTime? = _repo.GetLastImportDate()
            If lastImport.HasValue Then
                lblStatusLastImport.Text = "最終取り込み: " & lastImport.Value.ToString("yyyy/MM/dd HH:mm")
            Else
                lblStatusLastImport.Text = "最終取り込み: -"
            End If
        Catch
            lblStatusLastImport.Text = "最終取り込み: -"
        End Try
        UpdateFolderCounts()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  メール一覧 ソート・列設定
    ' ════════════════════════════════════════════════════════════

    ''' <summary>列名クリックでソート方向を切り替え、一覧を更新する。</summary>
    Private Sub listViewEmails_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles listViewEmails.ColumnClick
        If e.Column = _sortColumn Then
            _sortAscending = Not _sortAscending
        Else
            _sortColumn = e.Column
            ' 受信日時はデフォルト降順、他は昇順
            _sortAscending = (e.Column <> 3)
        End If
        UpdateColumnSortIndicator()
        SortEmailCache()
        listViewEmails.Invalidate()
    End Sub

    ''' <summary>_emailCache を現在のソート設定で並び替える。</summary>
    Private Sub SortEmailCache()
        If _emailCache Is Nothing OrElse _emailCache.Count = 0 Then Return
        Dim asc As Boolean = _sortAscending
        Select Case _sortColumn
            Case 0  ' 添付
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = a.HasAttachments.CompareTo(b.HasAttachments)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case 1  ' 件名
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = String.Compare(
                                         If(a.Subject, String.Empty),
                                         If(b.Subject, String.Empty),
                                         StringComparison.OrdinalIgnoreCase)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case 2  ' 差出人
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
            Case 3  ' 受信日時
                _emailCache.Sort(Function(a As Models.Email, b As Models.Email) As Integer
                                     Dim cmp As Integer = a.ReceivedAt.CompareTo(b.ReceivedAt)
                                     Return If(asc, cmp, -cmp)
                                 End Function)
            Case 4  ' サイズ
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
            Dim indicator As String = If(i = _sortColumn,
                                         If(_sortAscending, " ↑", " ↓"),
                                         String.Empty)
            listViewEmails.Columns(i).Text = _colBaseNames(i) & indicator
        Next
    End Sub

    ''' <summary>AppSettings から列幅を復元する。</summary>
    Private Sub RestoreColumnWidths()
        Dim colWidths As String = _settings.EmailListColumnWidths
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
        Dim colOrder As String = _settings.EmailListColumnOrder
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
        _settings.EmailListColumnOrder = String.Join(",", displayIndices)
        _settings.EmailListColumnWidths = String.Join(",", widths)
        _settings.EmailListSortColumn = _sortColumn
        _settings.EmailListSortAscending = _sortAscending
    End Sub

    ''' <summary>フォーム終了時に列設定を保存する。</summary>
    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SaveColumnSettings()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  メニューハンドラ
    ' ════════════════════════════════════════════════════════════

    Private Sub menuItemFileExit_Click(sender As Object, e As EventArgs) Handles menuItemFileExit.Click
        Me.Close()
    End Sub

    Private Sub menuItemSettings_Click(sender As Object, e As EventArgs) Handles menuItemSettings.Click
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
    Private Sub ReinitializeApp()
        ' サービス再生成（DB ファイルを新規作成してスキーマを初期化）
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
        LoadEmails(Nothing)
        UpdateStatusBar()

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
        End If
    End Sub

    ' ── テーブルビューア ──────────────────────────────────────

    Private Sub menuItemTableEmails_Click(sender As Object, e As EventArgs) Handles menuItemTableEmails.Click
        ShowTableViewer("emails")
    End Sub

    Private Sub menuItemTableAttachments_Click(sender As Object, e As EventArgs) Handles menuItemTableAttachments.Click
        ShowTableViewer("attachments")
    End Sub

    Private Sub menuItemTableDeletedIds_Click(sender As Object, e As EventArgs) Handles menuItemTableDeletedIds.Click
        ShowTableViewer("deleted_message_ids")
    End Sub

    Private Sub menuItemTableExchangeCache_Click(sender As Object, e As EventArgs) Handles menuItemTableExchangeCache.Click
        ShowTableViewer("exchange_address_cache")
    End Sub

    Private Sub ShowTableViewer(tableName As String)
        Using frm As New Forms.TableViewerForm(_dbManager, tableName)
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub menuItemHelpAbout_Click(sender As Object, e As EventArgs) Handles menuItemHelpAbout.Click
        MessageBox.Show("OutlookArchiver" & vbCrLf & "Outlook メールアーカイブツール",
            "バージョン情報", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
