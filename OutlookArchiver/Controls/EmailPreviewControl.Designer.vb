Option Explicit On
Option Strict On
Option Infer Off

Namespace Controls

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class EmailPreviewControl

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
            Me.pnlHeader = New System.Windows.Forms.Panel()
            Me.tlpHeader = New System.Windows.Forms.TableLayoutPanel()
            Me.lblFromCaption = New System.Windows.Forms.Label()
            Me.lblFromValue = New System.Windows.Forms.Label()
            Me.lblDateCaption = New System.Windows.Forms.Label()
            Me.lblDateValue = New System.Windows.Forms.Label()
            Me.lblSubjectCaption = New System.Windows.Forms.Label()
            Me.lblSubjectValue = New System.Windows.Forms.Label()
            Me.lblToCaption = New System.Windows.Forms.Label()
            Me.lblToValue = New System.Windows.Forms.Label()

            Me.pnlBody = New System.Windows.Forms.Panel()
            Me.webBrowser = New System.Windows.Forms.WebBrowser()
            Me.txtBodyText = New System.Windows.Forms.RichTextBox()

            Me.pnlAttachments = New System.Windows.Forms.Panel()
            Me.lblAttachTitle = New System.Windows.Forms.Label()
            Me.flowAttachments = New System.Windows.Forms.FlowLayoutPanel()

            ' ── Begin Init ─────────────────────────────────────────────
            Me.pnlHeader.SuspendLayout()
            Me.tlpHeader.SuspendLayout()
            Me.pnlBody.SuspendLayout()
            Me.pnlAttachments.SuspendLayout()
            Me.SuspendLayout()

            ' ── pnlHeader (Dock=Top, AutoSize) ──────────────────────────
            Me.pnlHeader.BackColor = System.Drawing.SystemColors.Info
            Me.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top
            Me.pnlHeader.AutoSize = True
            Me.pnlHeader.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Me.pnlHeader.MinimumSize = New System.Drawing.Size(0, 88)
            Me.pnlHeader.Name = "pnlHeader"
            Me.pnlHeader.Padding = New System.Windows.Forms.Padding(4, 2, 4, 2)
            Me.pnlHeader.Controls.Add(Me.tlpHeader)

            ' ── tlpHeader (2列・4行) ─────────────────────────────────────
            Me.tlpHeader.Dock = System.Windows.Forms.DockStyle.Fill
            Me.tlpHeader.AutoSize = True
            Me.tlpHeader.Name = "tlpHeader"
            Me.tlpHeader.ColumnCount = 2
            Me.tlpHeader.RowCount = 4
            Me.tlpHeader.ColumnStyles.Add(
                New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
            Me.tlpHeader.ColumnStyles.Add(
                New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
            Me.tlpHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22.0!))
            Me.tlpHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22.0!))
            Me.tlpHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22.0!))
            Me.tlpHeader.RowStyles.Add(
                New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))

            ' キャプションラベル
            Me.lblFromCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblFromCaption.Name = "lblFromCaption"
            Me.lblFromCaption.Text = "差出人:"
            Me.lblFromCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            Me.lblDateCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblDateCaption.Name = "lblDateCaption"
            Me.lblDateCaption.Text = "日時:"
            Me.lblDateCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            Me.lblSubjectCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblSubjectCaption.Name = "lblSubjectCaption"
            Me.lblSubjectCaption.Text = "件名:"
            Me.lblSubjectCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            Me.lblToCaption.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblToCaption.Name = "lblToCaption"
            Me.lblToCaption.Text = "宛先:"
            Me.lblToCaption.TextAlign = System.Drawing.ContentAlignment.MiddleRight

            ' 値ラベル
            Me.lblFromValue.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblFromValue.Name = "lblFromValue"
            Me.lblFromValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblFromValue.AutoEllipsis = True

            Me.lblDateValue.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblDateValue.Name = "lblDateValue"
            Me.lblDateValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

            Me.lblSubjectValue.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblSubjectValue.Name = "lblSubjectValue"
            Me.lblSubjectValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            Me.lblSubjectValue.AutoEllipsis = True

            Me.lblToValue.Dock = System.Windows.Forms.DockStyle.Fill
            Me.lblToValue.Name = "lblToValue"
            Me.lblToValue.TextAlign = System.Drawing.ContentAlignment.TopLeft
            Me.lblToValue.AutoSize = True
            Me.lblToValue.MaximumSize = New System.Drawing.Size(0, 0)
            Me.lblToValue.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)

            Me.tlpHeader.Controls.Add(Me.lblFromCaption, 0, 0)
            Me.tlpHeader.Controls.Add(Me.lblFromValue, 1, 0)
            Me.tlpHeader.Controls.Add(Me.lblDateCaption, 0, 1)
            Me.tlpHeader.Controls.Add(Me.lblDateValue, 1, 1)
            Me.tlpHeader.Controls.Add(Me.lblSubjectCaption, 0, 2)
            Me.tlpHeader.Controls.Add(Me.lblSubjectValue, 1, 2)
            Me.tlpHeader.Controls.Add(Me.lblToCaption, 0, 3)
            Me.tlpHeader.Controls.Add(Me.lblToValue, 1, 3)

            ' ── pnlAttachments (Dock=Bottom, 100px) ────────────────────
            Me.pnlAttachments.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.pnlAttachments.Height = 110
            Me.pnlAttachments.Name = "pnlAttachments"
            Me.pnlAttachments.Padding = New System.Windows.Forms.Padding(4)
            Me.pnlAttachments.Visible = False
            Me.pnlAttachments.Controls.Add(Me.flowAttachments)
            Me.pnlAttachments.Controls.Add(Me.lblAttachTitle)

            ' lblAttachTitle
            Me.lblAttachTitle.Dock = System.Windows.Forms.DockStyle.Top
            Me.lblAttachTitle.Name = "lblAttachTitle"
            Me.lblAttachTitle.Text = "添付ファイル:"
            Me.lblAttachTitle.Height = 22
            Me.lblAttachTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

            ' flowAttachments
            Me.flowAttachments.Dock = System.Windows.Forms.DockStyle.Fill
            Me.flowAttachments.Name = "flowAttachments"
            Me.flowAttachments.AutoScroll = True
            Me.flowAttachments.WrapContents = True
            Me.flowAttachments.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight

            ' ── pnlBody (Dock=Fill) ─────────────────────────────────────
            Me.pnlBody.Dock = System.Windows.Forms.DockStyle.Fill
            Me.pnlBody.Name = "pnlBody"
            Me.pnlBody.Controls.Add(Me.webBrowser)
            Me.pnlBody.Controls.Add(Me.txtBodyText)

            ' webBrowser
            Me.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill
            Me.webBrowser.Name = "webBrowser"
            Me.webBrowser.IsWebBrowserContextMenuEnabled = False
            Me.webBrowser.WebBrowserShortcutsEnabled = False
            Me.webBrowser.ScriptErrorsSuppressed = True

            ' txtBodyText
            Me.txtBodyText.Dock = System.Windows.Forms.DockStyle.Fill
            Me.txtBodyText.Name = "txtBodyText"
            Me.txtBodyText.ReadOnly = True
            Me.txtBodyText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
            Me.txtBodyText.WordWrap = True
            Me.txtBodyText.BackColor = System.Drawing.SystemColors.Window
            Me.txtBodyText.Visible = False

            ' ── UserControl 設定 ─────────────────────────────────────────
            ' 追加順序: pnlBody → pnlAttachments → pnlHeader
            '  (後から追加したものが先にDock処理される)
            Me.Controls.Add(Me.pnlBody)
            Me.Controls.Add(Me.pnlAttachments)
            Me.Controls.Add(Me.pnlHeader)
            Me.Name = "EmailPreviewControl"

            ' ── End Init ───────────────────────────────────────────────
            Me.pnlAttachments.ResumeLayout(False)
            Me.pnlBody.ResumeLayout(False)
            Me.tlpHeader.ResumeLayout(False)
            Me.pnlHeader.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

        ' ── コントロール宣言 ────────────────────────────────────────────
        Friend WithEvents pnlHeader As System.Windows.Forms.Panel
        Friend WithEvents tlpHeader As System.Windows.Forms.TableLayoutPanel
        Friend WithEvents lblFromCaption As System.Windows.Forms.Label
        Friend WithEvents lblFromValue As System.Windows.Forms.Label
        Friend WithEvents lblDateCaption As System.Windows.Forms.Label
        Friend WithEvents lblDateValue As System.Windows.Forms.Label
        Friend WithEvents lblSubjectCaption As System.Windows.Forms.Label
        Friend WithEvents lblSubjectValue As System.Windows.Forms.Label
        Friend WithEvents lblToCaption As System.Windows.Forms.Label
        Friend WithEvents lblToValue As System.Windows.Forms.Label
        Friend WithEvents pnlBody As System.Windows.Forms.Panel
        Friend WithEvents webBrowser As System.Windows.Forms.WebBrowser
        Friend WithEvents txtBodyText As System.Windows.Forms.RichTextBox
        Friend WithEvents pnlAttachments As System.Windows.Forms.Panel
        Friend WithEvents lblAttachTitle As System.Windows.Forms.Label
        Friend WithEvents flowAttachments As System.Windows.Forms.FlowLayoutPanel

    End Class

End Namespace
