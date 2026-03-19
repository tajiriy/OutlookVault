Option Explicit On
Option Strict On
Option Infer Off

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.menuStrip = New System.Windows.Forms.MenuStrip()
        Me.menuItemFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemSettings = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemFileSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuItemFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemImportNow = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemImportCancel = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemErrorExclusion = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemDev = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemDevTableViewer = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemDevAttachmentStats = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemHelpManual = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStrip = New System.Windows.Forms.ToolStrip()
        Me.btnImportNow = New System.Windows.Forms.ToolStripButton()
        Me.btnAutoImport = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.txtSearch = New System.Windows.Forms.ToolStripTextBox()
        Me.btnSearch = New System.Windows.Forms.ToolStripButton()
        Me.btnClearSearch = New System.Windows.Forms.ToolStripButton()
        Me.lblFolderCount = New System.Windows.Forms.ToolStripLabel()
        Me.btnSettings = New System.Windows.Forms.ToolStripButton()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.treeViewFolders = New System.Windows.Forms.TreeView()
        Me.splitRight = New System.Windows.Forms.SplitContainer()
        Me.listViewEmails = New System.Windows.Forms.ListView()
        Me.colAttach = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colSubject = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colSender = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colReceivedAt = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colSize = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.listViewContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.deleteMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.restoreMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.purgeMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.folderContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.emptyTrashMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.emailImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.tabControl = New System.Windows.Forms.TabControl()
        Me.tabPageNormal = New System.Windows.Forms.TabPage()
        Me.emailPreview = New OutlookVault.Controls.EmailPreviewControl()
        Me.tabPageThread = New System.Windows.Forms.TabPage()
        Me.conversationView = New OutlookVault.Controls.ConversationViewControl()
        Me.btnToggleView = New System.Windows.Forms.Button()
        Me.statusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblStatusCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblStatusSep = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblStatusLastImport = New System.Windows.Forms.ToolStripStatusLabel()
        Me.notifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.trayContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.trayMenuShow = New System.Windows.Forms.ToolStripMenuItem()
        Me.trayMenuImportNow = New System.Windows.Forms.ToolStripMenuItem()
        Me.trayMenuSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.trayMenuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me._autoImportTimer = New System.Windows.Forms.Timer(Me.components)
        Me._scheduledImportTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuStrip.SuspendLayout()
        Me.toolStrip.SuspendLayout()
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitMain.Panel1.SuspendLayout()
        Me.splitMain.Panel2.SuspendLayout()
        Me.splitMain.SuspendLayout()
        CType(Me.splitRight, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitRight.Panel1.SuspendLayout()
        Me.splitRight.Panel2.SuspendLayout()
        Me.splitRight.SuspendLayout()
        Me.listViewContextMenu.SuspendLayout()
        Me.folderContextMenu.SuspendLayout()
        Me.tabControl.SuspendLayout()
        Me.tabPageNormal.SuspendLayout()
        Me.tabPageThread.SuspendLayout()
        Me.statusStrip.SuspendLayout()
        Me.trayContextMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuStrip
        '
        Me.menuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemFile, Me.menuItemImport, Me.menuItemDev, Me.menuItemHelp})
        Me.menuStrip.Location = New System.Drawing.Point(0, 0)
        Me.menuStrip.Name = "menuStrip"
        Me.menuStrip.Size = New System.Drawing.Size(1397, 24)
        Me.menuStrip.TabIndex = 0
        '
        'menuItemFile
        '
        Me.menuItemFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemSettings, Me.menuItemFileSeparator1, Me.menuItemFileExit})
        Me.menuItemFile.Name = "menuItemFile"
        Me.menuItemFile.Size = New System.Drawing.Size(67, 20)
        Me.menuItemFile.Text = "ファイル(&F)"
        '
        'menuItemSettings
        '
        Me.menuItemSettings.Name = "menuItemSettings"
        Me.menuItemSettings.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Oemcomma), System.Windows.Forms.Keys)
        Me.menuItemSettings.ShortcutKeyDisplayString = "Ctrl+,"
        Me.menuItemSettings.Size = New System.Drawing.Size(155, 22)
        Me.menuItemSettings.Text = "設定(&T)..."
        '
        'menuItemFileSeparator1
        '
        Me.menuItemFileSeparator1.Name = "menuItemFileSeparator1"
        Me.menuItemFileSeparator1.Size = New System.Drawing.Size(152, 6)
        '
        'menuItemFileExit
        '
        Me.menuItemFileExit.Name = "menuItemFileExit"
        Me.menuItemFileExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.menuItemFileExit.Size = New System.Drawing.Size(155, 22)
        Me.menuItemFileExit.Text = "終了(&X)"
        '
        'menuItemImport
        '
        Me.menuItemImport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemImportNow, Me.menuItemImportCancel, Me.menuItemErrorExclusion})
        Me.menuItemImport.Name = "menuItemImport"
        Me.menuItemImport.Size = New System.Drawing.Size(73, 20)
        Me.menuItemImport.Text = "取り込み(&I)"
        '
        'menuItemImportNow
        '
        Me.menuItemImportNow.Name = "menuItemImportNow"
        Me.menuItemImportNow.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.menuItemImportNow.Size = New System.Drawing.Size(183, 22)
        Me.menuItemImportNow.Text = "今すぐ取り込み(&N)"
        '
        'menuItemImportCancel
        '
        Me.menuItemImportCancel.Enabled = False
        Me.menuItemImportCancel.Name = "menuItemImportCancel"
        Me.menuItemImportCancel.Size = New System.Drawing.Size(183, 22)
        Me.menuItemImportCancel.Text = "取り込み中止(&C)"
        '
        'menuItemErrorExclusion
        '
        Me.menuItemErrorExclusion.Name = "menuItemErrorExclusion"
        Me.menuItemErrorExclusion.Size = New System.Drawing.Size(183, 22)
        Me.menuItemErrorExclusion.Text = "エラー除外リスト(&E)..."
        '
        'menuItemDev
        '
        Me.menuItemDev.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemDevTableViewer, Me.menuItemDevAttachmentStats})
        Me.menuItemDev.Name = "menuItemDev"
        Me.menuItemDev.Size = New System.Drawing.Size(59, 20)
        Me.menuItemDev.Text = "開発(&D)"
        '
        'menuItemDevTableViewer
        '
        Me.menuItemDevTableViewer.Name = "menuItemDevTableViewer"
        Me.menuItemDevTableViewer.Size = New System.Drawing.Size(199, 22)
        Me.menuItemDevTableViewer.Text = "テーブルビューア(&T)"
        '
        'menuItemDevAttachmentStats
        '
        Me.menuItemDevAttachmentStats.Name = "menuItemDevAttachmentStats"
        Me.menuItemDevAttachmentStats.Size = New System.Drawing.Size(199, 22)
        Me.menuItemDevAttachmentStats.Text = "添付ファイル拡張子統計(&A)..."
        '
        'menuItemHelp
        '
        Me.menuItemHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemHelpManual, Me.menuItemHelpAbout})
        Me.menuItemHelp.Name = "menuItemHelp"
        Me.menuItemHelp.Size = New System.Drawing.Size(65, 20)
        Me.menuItemHelp.Text = "ヘルプ(&H)"
        '
        'menuItemHelpManual
        '
        Me.menuItemHelpManual.Name = "menuItemHelpManual"
        Me.menuItemHelpManual.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.menuItemHelpManual.Size = New System.Drawing.Size(193, 22)
        Me.menuItemHelpManual.Text = "ユーザーマニュアル(&M)"
        '
        'menuItemHelpAbout
        '
        Me.menuItemHelpAbout.Name = "menuItemHelpAbout"
        Me.menuItemHelpAbout.Size = New System.Drawing.Size(193, 22)
        Me.menuItemHelpAbout.Text = "バージョン情報(&A)"
        '
        'toolStrip
        '
        Me.toolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnImportNow, Me.btnAutoImport, Me.toolStripSep1, Me.txtSearch, Me.btnSearch, Me.btnClearSearch, Me.btnSettings, Me.lblFolderCount})
        Me.toolStrip.Location = New System.Drawing.Point(0, 24)
        Me.toolStrip.Name = "toolStrip"
        Me.toolStrip.Size = New System.Drawing.Size(1397, 25)
        Me.toolStrip.TabIndex = 1
        '
        'btnImportNow
        '
        Me.btnImportNow.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnImportNow.Name = "btnImportNow"
        Me.btnImportNow.Size = New System.Drawing.Size(84, 22)
        Me.btnImportNow.Text = "今すぐ取り込み"
        Me.btnImportNow.ToolTipText = "対象フォルダのメールを今すぐ取り込みます (F5)"
        '
        'btnAutoImport
        '
        Me.btnAutoImport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnAutoImport.Name = "btnAutoImport"
        Me.btnAutoImport.Size = New System.Drawing.Size(53, 22)
        Me.btnAutoImport.Text = "自動: ▶"
        Me.btnAutoImport.ToolTipText = "自動取り込みの開始/停止"
        '
        'toolStripSep1
        '
        Me.toolStripSep1.Name = "toolStripSep1"
        Me.toolStripSep1.Size = New System.Drawing.Size(6, 25)
        '
        'txtSearch
        '
        Me.txtSearch.Font = New System.Drawing.Font("Yu Gothic UI", 9.0!)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(200, 25)
        Me.txtSearch.ToolTipText = "検索キーワードを入力してEnterまたは検索ボタンを押してください"
        '
        'btnSearch
        '
        Me.btnSearch.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(35, 22)
        Me.btnSearch.Text = "検索"
        Me.btnSearch.ToolTipText = "メールを検索します"
        '
        'btnClearSearch
        '
        Me.btnClearSearch.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnClearSearch.Name = "btnClearSearch"
        Me.btnClearSearch.Size = New System.Drawing.Size(37, 22)
        Me.btnClearSearch.Text = "クリア"
        Me.btnClearSearch.ToolTipText = "検索をクリアします"
        '
        'lblFolderCount
        '
        Me.lblFolderCount.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.lblFolderCount.Name = "lblFolderCount"
        Me.lblFolderCount.Size = New System.Drawing.Size(0, 22)
        '
        'btnSettings
        '
        Me.btnSettings.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.btnSettings.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnSettings.Font = New System.Drawing.Font("Segoe UI Emoji", 12.0!)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(23, 22)
        Me.btnSettings.Text = "🔧"
        Me.btnSettings.ToolTipText = "設定を開きます (Ctrl+,)"
        '
        'splitMain
        '
        Me.splitMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitMain.Location = New System.Drawing.Point(0, 49)
        Me.splitMain.Name = "splitMain"
        '
        'splitMain.Panel1
        '
        Me.splitMain.Panel1.Controls.Add(Me.treeViewFolders)
        '
        'splitMain.Panel2
        '
        Me.splitMain.Panel2.Controls.Add(Me.splitRight)
        Me.splitMain.Size = New System.Drawing.Size(1397, 726)
        Me.splitMain.SplitterDistance = 233
        Me.splitMain.TabIndex = 0
        '
        'treeViewFolders
        '
        Me.treeViewFolders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeViewFolders.HideSelection = False
        Me.treeViewFolders.Location = New System.Drawing.Point(0, 0)
        Me.treeViewFolders.Name = "treeViewFolders"
        Me.treeViewFolders.Size = New System.Drawing.Size(233, 726)
        Me.treeViewFolders.TabIndex = 0
        '
        'splitRight
        '
        Me.splitRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitRight.Location = New System.Drawing.Point(0, 0)
        Me.splitRight.Name = "splitRight"
        Me.splitRight.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitRight.Panel1
        '
        Me.splitRight.Panel1.Controls.Add(Me.listViewEmails)
        '
        'splitRight.Panel2
        '
        Me.splitRight.Panel2.Controls.Add(Me.tabControl)
        Me.splitRight.Panel2.Controls.Add(Me.btnToggleView)
        Me.splitRight.Size = New System.Drawing.Size(1160, 726)
        Me.splitRight.SplitterDistance = 282
        Me.splitRight.TabIndex = 0
        '
        'listViewEmails
        '
        Me.listViewEmails.AllowColumnReorder = True
        Me.listViewEmails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colAttach, Me.colSubject, Me.colSender, Me.colReceivedAt, Me.colSize})
        Me.listViewEmails.ContextMenuStrip = Me.listViewContextMenu
        Me.listViewEmails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.listViewEmails.FullRowSelect = True
        Me.listViewEmails.MultiSelect = True
        Me.listViewEmails.HideSelection = False
        Me.listViewEmails.Location = New System.Drawing.Point(0, 0)
        Me.listViewEmails.Name = "listViewEmails"
        Me.listViewEmails.Size = New System.Drawing.Size(1160, 282)
        Me.listViewEmails.SmallImageList = Me.emailImageList
        Me.listViewEmails.TabIndex = 0
        Me.listViewEmails.UseCompatibleStateImageBehavior = False
        Me.listViewEmails.View = System.Windows.Forms.View.Details
        Me.listViewEmails.VirtualMode = True
        '
        'colAttach
        '
        Me.colAttach.Text = "添付"
        Me.colAttach.Width = 30
        '
        'colSubject
        '
        Me.colSubject.Text = "件名"
        Me.colSubject.Width = 720
        '
        'colSender
        '
        Me.colSender.Text = "差出人"
        Me.colSender.Width = 160
        '
        'colReceivedAt
        '
        Me.colReceivedAt.Text = "受信日時"
        Me.colReceivedAt.Width = 140
        '
        'colSize
        '
        Me.colSize.Text = "サイズ"
        Me.colSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.colSize.Width = 80
        '
        'listViewContextMenu
        '
        Me.listViewContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.deleteMenuItem})
        Me.listViewContextMenu.Name = "listViewContextMenu"
        Me.listViewContextMenu.Size = New System.Drawing.Size(115, 26)
        '
        'deleteMenuItem
        '
        Me.deleteMenuItem.Name = "deleteMenuItem"
        Me.deleteMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.deleteMenuItem.Text = "削除(&D)"
        '
        'restoreMenuItem
        '
        Me.restoreMenuItem.Name = "restoreMenuItem"
        Me.restoreMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.restoreMenuItem.Text = "復元(&R)"
        '
        'purgeMenuItem
        '
        Me.purgeMenuItem.Name = "purgeMenuItem"
        Me.purgeMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.purgeMenuItem.Text = "完全に削除(&P)"
        '
        'folderContextMenu
        '
        Me.folderContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.emptyTrashMenuItem})
        Me.folderContextMenu.Name = "folderContextMenu"
        Me.folderContextMenu.Size = New System.Drawing.Size(180, 26)
        '
        'emptyTrashMenuItem
        '
        Me.emptyTrashMenuItem.Name = "emptyTrashMenuItem"
        Me.emptyTrashMenuItem.Size = New System.Drawing.Size(179, 22)
        Me.emptyTrashMenuItem.Text = "ゴミ箱を空にする(&E)"
        '
        'emailImageList
        '
        Me.emailImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.emailImageList.ImageSize = New System.Drawing.Size(16, 16)
        Me.emailImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'tabControl
        '
        Me.tabControl.Controls.Add(Me.tabPageNormal)
        Me.tabControl.Controls.Add(Me.tabPageThread)
        Me.tabControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabControl.Location = New System.Drawing.Point(0, 0)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(1160, 440)
        Me.tabControl.TabIndex = 0
        '
        'tabPageNormal
        '
        Me.tabPageNormal.Controls.Add(Me.emailPreview)
        Me.tabPageNormal.Location = New System.Drawing.Point(4, 24)
        Me.tabPageNormal.Name = "tabPageNormal"
        Me.tabPageNormal.Size = New System.Drawing.Size(1152, 412)
        Me.tabPageNormal.TabIndex = 0
        Me.tabPageNormal.Text = "通常表示"
        Me.tabPageNormal.UseVisualStyleBackColor = True
        '
        'emailPreview
        '
        Me.emailPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.emailPreview.Location = New System.Drawing.Point(0, 0)
        Me.emailPreview.Name = "emailPreview"
        Me.emailPreview.Size = New System.Drawing.Size(1152, 412)
        Me.emailPreview.TabIndex = 0
        '
        'tabPageThread
        '
        Me.tabPageThread.Controls.Add(Me.conversationView)
        Me.tabPageThread.Location = New System.Drawing.Point(4, 22)
        Me.tabPageThread.Name = "tabPageThread"
        Me.tabPageThread.Size = New System.Drawing.Size(1146, 414)
        Me.tabPageThread.TabIndex = 1
        Me.tabPageThread.Text = "会話ビュー"
        Me.tabPageThread.UseVisualStyleBackColor = True
        '
        'conversationView
        '
        Me.conversationView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.conversationView.Location = New System.Drawing.Point(0, 0)
        Me.conversationView.Name = "conversationView"
        Me.conversationView.Size = New System.Drawing.Size(1146, 414)
        Me.conversationView.TabIndex = 0
        '
        'btnToggleView
        '
        Me.btnToggleView.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnToggleView.AutoSize = True
        Me.btnToggleView.Enabled = False
        Me.btnToggleView.Location = New System.Drawing.Point(41, 0)
        Me.btnToggleView.Name = "btnToggleView"
        Me.btnToggleView.Size = New System.Drawing.Size(77, 25)
        Me.btnToggleView.TabIndex = 1
        Me.btnToggleView.Text = "テキスト表示"
        '
        'statusStrip
        '
        Me.statusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatusCount, Me.lblStatusSep, Me.lblStatusLastImport})
        Me.statusStrip.Location = New System.Drawing.Point(0, 775)
        Me.statusStrip.Name = "statusStrip"
        Me.statusStrip.Size = New System.Drawing.Size(1397, 22)
        Me.statusStrip.TabIndex = 3
        '
        'lblStatusCount
        '
        Me.lblStatusCount.Name = "lblStatusCount"
        Me.lblStatusCount.Size = New System.Drawing.Size(52, 17)
        Me.lblStatusCount.Text = "総数 0件"
        '
        'lblStatusSep
        '
        Me.lblStatusSep.Name = "lblStatusSep"
        Me.lblStatusSep.Size = New System.Drawing.Size(16, 17)
        Me.lblStatusSep.Text = " | "
        '
        'lblStatusLastImport
        '
        Me.lblStatusLastImport.Name = "lblStatusLastImport"
        Me.lblStatusLastImport.Size = New System.Drawing.Size(85, 17)
        Me.lblStatusLastImport.Text = "最終取り込み: -"
        '
        'notifyIcon
        '
        Me.notifyIcon.ContextMenuStrip = Me.trayContextMenu
        Me.notifyIcon.Text = "OutlookVault"
        '
        'trayContextMenu
        '
        Me.trayContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.trayMenuShow, Me.trayMenuImportNow, Me.trayMenuSep1, Me.trayMenuExit})
        Me.trayContextMenu.Name = "trayContextMenu"
        Me.trayContextMenu.Size = New System.Drawing.Size(159, 76)
        '
        'trayMenuShow
        '
        Me.trayMenuShow.Font = New System.Drawing.Font("Meiryo UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.trayMenuShow.Name = "trayMenuShow"
        Me.trayMenuShow.Size = New System.Drawing.Size(158, 22)
        Me.trayMenuShow.Text = "表示(&S)"
        '
        'trayMenuImportNow
        '
        Me.trayMenuImportNow.Name = "trayMenuImportNow"
        Me.trayMenuImportNow.Size = New System.Drawing.Size(158, 22)
        Me.trayMenuImportNow.Text = "今すぐ取り込み(&I)"
        '
        'trayMenuSep1
        '
        Me.trayMenuSep1.Name = "trayMenuSep1"
        Me.trayMenuSep1.Size = New System.Drawing.Size(155, 6)
        '
        'trayMenuExit
        '
        Me.trayMenuExit.Name = "trayMenuExit"
        Me.trayMenuExit.Size = New System.Drawing.Size(158, 22)
        Me.trayMenuExit.Text = "終了(&X)"
        '
        '_autoImportTimer
        '
        Me._autoImportTimer.Interval = 600000
        '
        '_scheduledImportTimer
        '
        Me._scheduledImportTimer.Interval = 60000
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1397, 797)
        Me.Controls.Add(Me.splitMain)
        Me.Controls.Add(Me.toolStrip)
        Me.Controls.Add(Me.menuStrip)
        Me.Controls.Add(Me.statusStrip)
        Me.Font = New System.Drawing.Font("Meiryo UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.menuStrip
        Me.Name = "MainForm"
        Me.Text = "OutlookVault"
        Me.menuStrip.ResumeLayout(False)
        Me.menuStrip.PerformLayout()
        Me.toolStrip.ResumeLayout(False)
        Me.toolStrip.PerformLayout()
        Me.splitMain.Panel1.ResumeLayout(False)
        Me.splitMain.Panel2.ResumeLayout(False)
        CType(Me.splitMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitMain.ResumeLayout(False)
        Me.splitRight.Panel1.ResumeLayout(False)
        Me.splitRight.Panel2.ResumeLayout(False)
        Me.splitRight.Panel2.PerformLayout()
        CType(Me.splitRight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitRight.ResumeLayout(False)
        Me.listViewContextMenu.ResumeLayout(False)
        Me.folderContextMenu.ResumeLayout(False)
        Me.tabControl.ResumeLayout(False)
        Me.tabPageNormal.ResumeLayout(False)
        Me.tabPageThread.ResumeLayout(False)
        Me.statusStrip.ResumeLayout(False)
        Me.statusStrip.PerformLayout()
        Me.trayContextMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    ' ── コントロール宣言 ────────────────────────────────────────────
    Friend WithEvents menuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents menuItemFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemFileSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents menuItemImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemImportNow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemImportCancel As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemErrorExclusion As System.Windows.Forms.ToolStripMenuItem

    Friend WithEvents menuItemSettings As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemDev As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemDevTableViewer As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemDevAttachmentStats As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemHelpManual As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents btnImportNow As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnAutoImport As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents txtSearch As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents btnSearch As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnClearSearch As System.Windows.Forms.ToolStripButton
    Friend WithEvents lblFolderCount As System.Windows.Forms.ToolStripLabel
    Friend WithEvents btnSettings As System.Windows.Forms.ToolStripButton
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents splitRight As System.Windows.Forms.SplitContainer
    Friend WithEvents treeViewFolders As System.Windows.Forms.TreeView
    Friend WithEvents listViewEmails As System.Windows.Forms.ListView
    Friend colAttach As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSubject As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSender As System.Windows.Forms.ColumnHeader
    Friend WithEvents colReceivedAt As System.Windows.Forms.ColumnHeader
    Friend colSize As System.Windows.Forms.ColumnHeader
    Friend emailImageList As System.Windows.Forms.ImageList
    Friend WithEvents listViewContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents deleteMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents restoreMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents purgeMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents folderContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents emptyTrashMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tabControl As System.Windows.Forms.TabControl
    Friend WithEvents tabPageNormal As System.Windows.Forms.TabPage
    Friend WithEvents tabPageThread As System.Windows.Forms.TabPage
    Friend WithEvents emailPreview As Controls.EmailPreviewControl
    Friend WithEvents conversationView As Controls.ConversationViewControl
    Friend WithEvents statusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents lblStatusCount As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblStatusSep As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblStatusLastImport As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnToggleView As System.Windows.Forms.Button
    Friend WithEvents notifyIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents trayContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents trayMenuShow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents trayMenuImportNow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents trayMenuSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents trayMenuExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents _autoImportTimer As System.Windows.Forms.Timer
    Friend WithEvents _scheduledImportTimer As System.Windows.Forms.Timer

End Class
