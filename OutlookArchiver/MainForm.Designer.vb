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
        Me.menuStrip = New System.Windows.Forms.MenuStrip()
        Me.menuItemFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemImportNow = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemSearch = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemSettings = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItemHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStrip = New System.Windows.Forms.ToolStrip()
        Me.btnImportNow = New System.Windows.Forms.ToolStripButton()
        Me.btnAutoImport = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSep1 = New System.Windows.Forms.ToolStripSeparator()
        Me.txtSearch = New System.Windows.Forms.ToolStripTextBox()
        Me.btnSearch = New System.Windows.Forms.ToolStripButton()
        Me.splitMain = New System.Windows.Forms.SplitContainer()
        Me.treeViewFolders = New System.Windows.Forms.TreeView()
        Me.splitRight = New System.Windows.Forms.SplitContainer()
        Me.listViewEmails = New System.Windows.Forms.ListView()
        Me.colSubject = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colSender = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.colReceivedAt = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.tabControl = New System.Windows.Forms.TabControl()
        Me.tabPageNormal = New System.Windows.Forms.TabPage()
        Me.tabPageThread = New System.Windows.Forms.TabPage()
        Me.statusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblStatusCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblStatusSep = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblStatusLastImport = New System.Windows.Forms.ToolStripStatusLabel()
        Me.emailPreview = New OutlookArchiver.Controls.EmailPreviewControl()
        Me.conversationView = New OutlookArchiver.Controls.ConversationViewControl()
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
        Me.tabControl.SuspendLayout()
        Me.tabPageNormal.SuspendLayout()
        Me.tabPageThread.SuspendLayout()
        Me.statusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuStrip
        '
        Me.menuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemFile, Me.menuItemImport, Me.menuItemSearch, Me.menuItemSettings, Me.menuItemHelp})
        Me.menuStrip.Location = New System.Drawing.Point(0, 0)
        Me.menuStrip.Name = "menuStrip"
        Me.menuStrip.Size = New System.Drawing.Size(1271, 24)
        Me.menuStrip.TabIndex = 0
        '
        'menuItemFile
        '
        Me.menuItemFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemFileExit})
        Me.menuItemFile.Name = "menuItemFile"
        Me.menuItemFile.Size = New System.Drawing.Size(67, 20)
        Me.menuItemFile.Text = "ファイル(&F)"
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
        Me.menuItemImport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemImportNow})
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
        'menuItemSearch
        '
        Me.menuItemSearch.Name = "menuItemSearch"
        Me.menuItemSearch.Size = New System.Drawing.Size(57, 20)
        Me.menuItemSearch.Text = "検索(&S)"
        '
        'menuItemSettings
        '
        Me.menuItemSettings.Name = "menuItemSettings"
        Me.menuItemSettings.Size = New System.Drawing.Size(57, 20)
        Me.menuItemSettings.Text = "設定(&T)"
        '
        'menuItemHelp
        '
        Me.menuItemHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemHelpAbout})
        Me.menuItemHelp.Name = "menuItemHelp"
        Me.menuItemHelp.Size = New System.Drawing.Size(65, 20)
        Me.menuItemHelp.Text = "ヘルプ(&H)"
        '
        'menuItemHelpAbout
        '
        Me.menuItemHelpAbout.Name = "menuItemHelpAbout"
        Me.menuItemHelpAbout.Size = New System.Drawing.Size(158, 22)
        Me.menuItemHelpAbout.Text = "バージョン情報(&A)"
        '
        'toolStrip
        '
        Me.toolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnImportNow, Me.btnAutoImport, Me.toolStripSep1, Me.txtSearch, Me.btnSearch})
        Me.toolStrip.Location = New System.Drawing.Point(0, 24)
        Me.toolStrip.Name = "toolStrip"
        Me.toolStrip.Size = New System.Drawing.Size(1271, 25)
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
        Me.splitMain.Size = New System.Drawing.Size(1271, 668)
        Me.splitMain.SplitterDistance = 150
        Me.splitMain.TabIndex = 0
        '
        'treeViewFolders
        '
        Me.treeViewFolders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeViewFolders.HideSelection = False
        Me.treeViewFolders.Location = New System.Drawing.Point(0, 0)
        Me.treeViewFolders.Name = "treeViewFolders"
        Me.treeViewFolders.Size = New System.Drawing.Size(150, 668)
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
        Me.splitRight.Size = New System.Drawing.Size(1117, 668)
        Me.splitRight.SplitterDistance = 260
        Me.splitRight.TabIndex = 0
        '
        'listViewEmails
        '
        Me.listViewEmails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colSubject, Me.colSender, Me.colReceivedAt})
        Me.listViewEmails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.listViewEmails.FullRowSelect = True
        Me.listViewEmails.HideSelection = False
        Me.listViewEmails.Location = New System.Drawing.Point(0, 0)
        Me.listViewEmails.MultiSelect = False
        Me.listViewEmails.Name = "listViewEmails"
        Me.listViewEmails.Size = New System.Drawing.Size(1117, 260)
        Me.listViewEmails.TabIndex = 0
        Me.listViewEmails.UseCompatibleStateImageBehavior = False
        Me.listViewEmails.View = System.Windows.Forms.View.Details
        Me.listViewEmails.VirtualMode = True
        '
        'colSubject
        '
        Me.colSubject.Text = "件名"
        Me.colSubject.Width = 360
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
        'tabControl
        '
        Me.tabControl.Controls.Add(Me.tabPageNormal)
        Me.tabControl.Controls.Add(Me.tabPageThread)
        Me.tabControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabControl.Location = New System.Drawing.Point(0, 0)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(1117, 404)
        Me.tabControl.TabIndex = 0
        '
        'tabPageNormal
        '
        Me.tabPageNormal.Controls.Add(Me.emailPreview)
        Me.tabPageNormal.Location = New System.Drawing.Point(4, 22)
        Me.tabPageNormal.Name = "tabPageNormal"
        Me.tabPageNormal.Size = New System.Drawing.Size(1109, 378)
        Me.tabPageNormal.TabIndex = 0
        Me.tabPageNormal.Text = "通常表示"
        Me.tabPageNormal.UseVisualStyleBackColor = True
        '
        'tabPageThread
        '
        Me.tabPageThread.Controls.Add(Me.conversationView)
        Me.tabPageThread.Location = New System.Drawing.Point(4, 22)
        Me.tabPageThread.Name = "tabPageThread"
        Me.tabPageThread.Size = New System.Drawing.Size(1109, 378)
        Me.tabPageThread.TabIndex = 1
        Me.tabPageThread.Text = "会話ビュー"
        Me.tabPageThread.UseVisualStyleBackColor = True
        '
        'conversationView
        '
        Me.conversationView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.conversationView.Location = New System.Drawing.Point(0, 0)
        Me.conversationView.Name = "conversationView"
        Me.conversationView.Size = New System.Drawing.Size(1109, 378)
        Me.conversationView.TabIndex = 0
        '
        'statusStrip
        '
        Me.statusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatusCount, Me.lblStatusSep, Me.lblStatusLastImport})
        Me.statusStrip.Location = New System.Drawing.Point(0, 717)
        Me.statusStrip.Name = "statusStrip"
        Me.statusStrip.Size = New System.Drawing.Size(1271, 22)
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
        'emailPreview
        '
        Me.emailPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.emailPreview.Location = New System.Drawing.Point(0, 0)
        Me.emailPreview.Name = "emailPreview"
        Me.emailPreview.Size = New System.Drawing.Size(1109, 378)
        Me.emailPreview.TabIndex = 0
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1271, 739)
        Me.Controls.Add(Me.splitMain)
        Me.Controls.Add(Me.toolStrip)
        Me.Controls.Add(Me.menuStrip)
        Me.Controls.Add(Me.statusStrip)
        Me.MainMenuStrip = Me.menuStrip
        Me.Name = "MainForm"
        Me.Text = "OutlookArchiver"
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
        CType(Me.splitRight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitRight.ResumeLayout(False)
        Me.tabControl.ResumeLayout(False)
        Me.tabPageNormal.ResumeLayout(False)
        Me.tabPageThread.ResumeLayout(False)
        Me.statusStrip.ResumeLayout(False)
        Me.statusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    ' ── コントロール宣言 ────────────────────────────────────────────
    Friend WithEvents menuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents menuItemFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemImportNow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemSearch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemSettings As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuItemHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents toolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents btnImportNow As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnAutoImport As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSep1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents txtSearch As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents btnSearch As System.Windows.Forms.ToolStripButton
    Friend WithEvents splitMain As System.Windows.Forms.SplitContainer
    Friend WithEvents splitRight As System.Windows.Forms.SplitContainer
    Friend WithEvents treeViewFolders As System.Windows.Forms.TreeView
    Friend WithEvents listViewEmails As System.Windows.Forms.ListView
    Friend WithEvents colSubject As System.Windows.Forms.ColumnHeader
    Friend WithEvents colSender As System.Windows.Forms.ColumnHeader
    Friend WithEvents colReceivedAt As System.Windows.Forms.ColumnHeader
    Friend WithEvents tabControl As System.Windows.Forms.TabControl
    Friend WithEvents tabPageNormal As System.Windows.Forms.TabPage
    Friend WithEvents tabPageThread As System.Windows.Forms.TabPage
    Friend WithEvents emailPreview As Controls.EmailPreviewControl
    Friend WithEvents conversationView As Controls.ConversationViewControl
    Friend WithEvents statusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents lblStatusCount As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblStatusSep As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblStatusLastImport As System.Windows.Forms.ToolStripStatusLabel

End Class
