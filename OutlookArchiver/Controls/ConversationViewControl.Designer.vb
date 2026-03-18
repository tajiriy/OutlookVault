Option Explicit On
Option Strict On
Option Infer Off

Namespace Controls

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ConversationViewControl

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        Private components As System.ComponentModel.IContainer

        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()

            ' ── コントロール生成 ────────────────────────────────────────
            Me.splitConversation = New System.Windows.Forms.SplitContainer()
            Me.listViewThread = New System.Windows.Forms.ListView()
            Me.colThreadSender = New System.Windows.Forms.ColumnHeader()
            Me.colThreadDate = New System.Windows.Forms.ColumnHeader()
            Me.colThreadSubject = New System.Windows.Forms.ColumnHeader()

            Me.pnlMsgHeader = New System.Windows.Forms.Panel()
            Me.tlpMsgHeader = New System.Windows.Forms.TableLayoutPanel()
            Me.lblMsgFromCaption = New System.Windows.Forms.Label()
            Me.lblMsgFrom = New System.Windows.Forms.Label()
            Me.lblMsgDateCaption = New System.Windows.Forms.Label()
            Me.lblMsgDate = New System.Windows.Forms.Label()

            Me.pnlMsgToolbar = New System.Windows.Forms.Panel()
            Me.btnToggleMsgView = New System.Windows.Forms.Button()

            Me.pnlMsgBody = New System.Windows.Forms.Panel()
            Me.webBrowserMsg = New System.Windows.Forms.WebBrowser()
            Me.txtBodyMsg = New System.Windows.Forms.RichTextBox()

            ' ── Begin Init ─────────────────────────────────────────────
            CType(Me.splitConversation, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitConversation.SuspendLayout()
            Me.pnlMsgHeader.SuspendLayout()
            Me.tlpMsgHeader.SuspendLayout()
            Me.pnlMsgToolbar.SuspendLayout()
            Me.pnlMsgBody.SuspendLayout()
            Me.SuspendLayout()

            ' ── splitConversation ────────────────────────────────────────
            Me.splitConversation.Dock = System.Windows.Forms.DockStyle.Fill
            Me.splitConversation.Name = "splitConversation"
            Me.splitConversation.Orientation = System.Windows.Forms.Orientation.Horizontal
            Me.splitConversation.SplitterDistance = 100
            Me.splitConversation.TabIndex = 0
            '
            ' Panel1: スレッド一覧
            '
            Me.splitConversation.Panel1.Controls.Add(Me.listViewThread)
            '
            ' Panel2: メッセージ本文エリア
            ' 追加順序: pnlMsgBody(Fill) → pnlMsgToolbar(Top) → pnlMsgHeader(Top)
            '  後から追加したものが先にDock処理される
            '
            Me.splitConversation.Panel2.Controls.Add(Me.pnlMsgBody)
            Me.splitConversation.Panel2.Controls.Add(Me.pnlMsgToolbar)
            Me.splitConversation.Panel2.Controls.Add(Me.pnlMsgHeader)

            ' ── listViewThread ───────────────────────────────────────────
            Me.listViewThread.Dock = System.Windows.Forms.DockStyle.Fill
            Me.listViewThread.Name = "listViewThread"
            Me.listViewThread.View = System.Windows.Forms.View.Details
            Me.listViewThread.FullRowSelect = True
            Me.listViewThread.HideSelection = False
            Me.listViewThread.MultiSelect = False
            Me.listViewThread.TabIndex = 0
            Me.listViewThread.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {
                Me.colThreadSender, Me.colThreadDate, Me.colThreadSubject})

            Me.colThreadSender.Text = "差出人"
            Me.colThreadSender.Width = 160

            Me.colThreadDate.Text = "日時"
            Me.colThreadDate.Width = 140

            Me.colThreadSubject.Text = "件名"
            Me.colThreadSubject.Width = 400

            ' ── pnlMsgHeader (Dock=Top, 50px) ───────────────────────────
            Me.pnlMsgHeader.BackColor = System.Drawing.SystemColors.Info
            Me.pnlMsgHeader.Dock = System.Windows.Forms.DockStyle.Top
            Me.pnlMsgHeader.Height = 50
            Me.pnlMsgHeader.Name = "pnlMsgHeader"
            Me.pnlMsgHeader.Padding = New System.Windows.Forms.Padding(4, 2, 4, 2)
            Me.pnlMsgHeader.Controls.Add(Me.tlpMsgHeader)

            ' ── tlpMsgHeader (2列・2行) ───────────────────────────────────
            Me.tlpMsgHeader.Dock = System.Windows.Forms.DockStyle.Fill
            Me.tlpMsgHeader.Name = "tlpMsgHeader"
            Me.tlpMsgHeader.ColumnCount = 2
            Me.tlpMsgHeader.RowCount = 2
            Me.tlpMsgHeader.ColumnStyles.Add(
                New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
            Me.tlpMsgHeader.ColumnStyles.Add(
                New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
            Me.tlpMsgHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
            Me.tlpMsgHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))

            Me.lblMsgFromCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblMsgFromCaption.Name = "lblMsgFromCaption"
            Me.lblMsgFromCaption.Text = "差出人:"
            Me.lblMsgFromCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            Me.lblMsgFrom.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblMsgFrom.Name = "lblMsgFrom"
            Me.lblMsgFrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblMsgFrom.AutoEllipsis = True

            Me.lblMsgDateCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblMsgDateCaption.Name = "lblMsgDateCaption"
            Me.lblMsgDateCaption.Text = "日時:"
            Me.lblMsgDateCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            Me.lblMsgDate.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblMsgDate.Name = "lblMsgDate"
            Me.lblMsgDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

            Me.tlpMsgHeader.Controls.Add(Me.lblMsgFromCaption, 0, 0)
            Me.tlpMsgHeader.Controls.Add(Me.lblMsgFrom, 1, 0)
            Me.tlpMsgHeader.Controls.Add(Me.lblMsgDateCaption, 0, 1)
            Me.tlpMsgHeader.Controls.Add(Me.lblMsgDate, 1, 1)

            ' ── pnlMsgToolbar (Dock=Top, 30px) ──────────────────────────
            Me.pnlMsgToolbar.Dock = System.Windows.Forms.DockStyle.Top
            Me.pnlMsgToolbar.Height = 30
            Me.pnlMsgToolbar.Name = "pnlMsgToolbar"
            Me.pnlMsgToolbar.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
            Me.pnlMsgToolbar.Controls.Add(Me.btnToggleMsgView)

            Me.btnToggleMsgView.Name = "btnToggleMsgView"
            Me.btnToggleMsgView.Text = "テキスト表示"
            Me.btnToggleMsgView.AutoSize = True
            Me.btnToggleMsgView.Height = 23
            Me.btnToggleMsgView.Enabled = False

            ' ── pnlMsgBody (Dock=Fill) ───────────────────────────────────
            Me.pnlMsgBody.Dock = System.Windows.Forms.DockStyle.Fill
            Me.pnlMsgBody.Name = "pnlMsgBody"
            Me.pnlMsgBody.Controls.Add(Me.webBrowserMsg)
            Me.pnlMsgBody.Controls.Add(Me.txtBodyMsg)

            ' webBrowserMsg
            Me.webBrowserMsg.Dock = System.Windows.Forms.DockStyle.Fill
            Me.webBrowserMsg.Name = "webBrowserMsg"
            Me.webBrowserMsg.IsWebBrowserContextMenuEnabled = False
            Me.webBrowserMsg.WebBrowserShortcutsEnabled = False
            Me.webBrowserMsg.ScriptErrorsSuppressed = True

            ' txtBodyMsg
            Me.txtBodyMsg.Dock = System.Windows.Forms.DockStyle.Fill
            Me.txtBodyMsg.Name = "txtBodyMsg"
            Me.txtBodyMsg.ReadOnly = True
            Me.txtBodyMsg.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
            Me.txtBodyMsg.WordWrap = True
            Me.txtBodyMsg.BackColor = System.Drawing.SystemColors.Window
            Me.txtBodyMsg.Visible = False

            ' ── UserControl 設定 ─────────────────────────────────────────
            Me.Controls.Add(Me.splitConversation)
            Me.Name = "ConversationViewControl"

            ' ── End Init ───────────────────────────────────────────────
            Me.pnlMsgBody.ResumeLayout(False)
            Me.pnlMsgToolbar.ResumeLayout(False)
            Me.tlpMsgHeader.ResumeLayout(False)
            Me.pnlMsgHeader.ResumeLayout(False)
            Me.splitConversation.Panel1.ResumeLayout(False)
            Me.splitConversation.Panel2.ResumeLayout(False)
            CType(Me.splitConversation, System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitConversation.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

        ' ── コントロール宣言 ────────────────────────────────────────────
        Friend WithEvents splitConversation As System.Windows.Forms.SplitContainer
        Friend WithEvents listViewThread As System.Windows.Forms.ListView
        Friend WithEvents colThreadSender As System.Windows.Forms.ColumnHeader
        Friend WithEvents colThreadDate As System.Windows.Forms.ColumnHeader
        Friend WithEvents colThreadSubject As System.Windows.Forms.ColumnHeader
        Friend WithEvents pnlMsgHeader As System.Windows.Forms.Panel
        Friend WithEvents tlpMsgHeader As System.Windows.Forms.TableLayoutPanel
        Friend WithEvents lblMsgFromCaption As System.Windows.Forms.Label
        Friend WithEvents lblMsgFrom As System.Windows.Forms.Label
        Friend WithEvents lblMsgDateCaption As System.Windows.Forms.Label
        Friend WithEvents lblMsgDate As System.Windows.Forms.Label
        Friend WithEvents pnlMsgToolbar As System.Windows.Forms.Panel
        Friend WithEvents btnToggleMsgView As System.Windows.Forms.Button
        Friend WithEvents pnlMsgBody As System.Windows.Forms.Panel
        Friend WithEvents webBrowserMsg As System.Windows.Forms.WebBrowser
        Friend WithEvents txtBodyMsg As System.Windows.Forms.RichTextBox

    End Class

End Namespace
