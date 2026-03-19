Option Explicit On
Option Strict On
Option Infer Off

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsForm
    Inherits System.Windows.Forms.Form

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

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.grpPaths = New System.Windows.Forms.GroupBox()
        Me.lblDbPath = New System.Windows.Forms.Label()
        Me.txtDbPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseDb = New System.Windows.Forms.Button()
        Me.lblAttachDir = New System.Windows.Forms.Label()
        Me.txtAttachDir = New System.Windows.Forms.TextBox()
        Me.btnBrowseAttach = New System.Windows.Forms.Button()
        Me.grpAutoImport = New System.Windows.Forms.GroupBox()
        Me.chkAutoImportEnabled = New System.Windows.Forms.CheckBox()
        Me.rdoInterval = New System.Windows.Forms.RadioButton()
        Me.rdoScheduled = New System.Windows.Forms.RadioButton()
        Me.numInterval = New System.Windows.Forms.NumericUpDown()
        Me.dtpScheduledTime = New System.Windows.Forms.DateTimePicker()
        Me.lblMaxCount = New System.Windows.Forms.Label()
        Me.numMaxCount = New System.Windows.Forms.NumericUpDown()
        Me.lblImportOrder = New System.Windows.Forms.Label()
        Me.cboImportOrder = New System.Windows.Forms.ComboBox()
        Me.grpFolders = New System.Windows.Forms.GroupBox()
        Me.lstFolders = New System.Windows.Forms.ListBox()
        Me.btnSelectFolders = New System.Windows.Forms.Button()
        Me.grpDisplay = New System.Windows.Forms.GroupBox()
        Me.chkDefaultHtml = New System.Windows.Forms.CheckBox()
        Me.chkSortAscending = New System.Windows.Forms.CheckBox()
        Me.chkShowImportResult = New System.Windows.Forms.CheckBox()
        Me.grpTray = New System.Windows.Forms.GroupBox()
        Me.chkMinimizeToTray = New System.Windows.Forms.CheckBox()
        Me.chkCloseToTray = New System.Windows.Forms.CheckBox()
        Me.chkShowBalloonOnImport = New System.Windows.Forms.CheckBox()
        Me.grpData = New System.Windows.Forms.GroupBox()
        Me.btnResetData = New System.Windows.Forms.Button()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpPaths.SuspendLayout()
        Me.grpAutoImport.SuspendLayout()
        CType(Me.numInterval, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numMaxCount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFolders.SuspendLayout()
        Me.grpDisplay.SuspendLayout()
        Me.grpTray.SuspendLayout()
        Me.grpData.SuspendLayout()
        Me.pnlButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpPaths
        '
        Me.grpPaths.Controls.Add(Me.lblDbPath)
        Me.grpPaths.Controls.Add(Me.txtDbPath)
        Me.grpPaths.Controls.Add(Me.btnBrowseDb)
        Me.grpPaths.Controls.Add(Me.lblAttachDir)
        Me.grpPaths.Controls.Add(Me.txtAttachDir)
        Me.grpPaths.Controls.Add(Me.btnBrowseAttach)
        Me.grpPaths.Location = New System.Drawing.Point(8, 8)
        Me.grpPaths.Name = "grpPaths"
        Me.grpPaths.Size = New System.Drawing.Size(488, 94)
        Me.grpPaths.TabIndex = 0
        Me.grpPaths.TabStop = False
        Me.grpPaths.Text = "保存先"
        '
        'lblDbPath
        '
        Me.lblDbPath.AutoSize = True
        Me.lblDbPath.Location = New System.Drawing.Point(8, 26)
        Me.lblDbPath.Name = "lblDbPath"
        Me.lblDbPath.Size = New System.Drawing.Size(75, 13)
        Me.lblDbPath.Text = "DB ファイル:"
        '
        'txtDbPath
        '
        Me.txtDbPath.Location = New System.Drawing.Point(132, 23)
        Me.txtDbPath.Name = "txtDbPath"
        Me.txtDbPath.Size = New System.Drawing.Size(256, 22)
        Me.txtDbPath.TabIndex = 1
        '
        'btnBrowseDb
        '
        Me.btnBrowseDb.Location = New System.Drawing.Point(396, 22)
        Me.btnBrowseDb.Name = "btnBrowseDb"
        Me.btnBrowseDb.Size = New System.Drawing.Size(80, 24)
        Me.btnBrowseDb.TabIndex = 2
        Me.btnBrowseDb.Text = "参照..."
        '
        'lblAttachDir
        '
        Me.lblAttachDir.AutoSize = True
        Me.lblAttachDir.Location = New System.Drawing.Point(8, 56)
        Me.lblAttachDir.Name = "lblAttachDir"
        Me.lblAttachDir.Size = New System.Drawing.Size(117, 13)
        Me.lblAttachDir.Text = "添付ファイル保存先:"
        '
        'txtAttachDir
        '
        Me.txtAttachDir.Location = New System.Drawing.Point(132, 53)
        Me.txtAttachDir.Name = "txtAttachDir"
        Me.txtAttachDir.Size = New System.Drawing.Size(256, 22)
        Me.txtAttachDir.TabIndex = 4
        '
        'btnBrowseAttach
        '
        Me.btnBrowseAttach.Location = New System.Drawing.Point(396, 52)
        Me.btnBrowseAttach.Name = "btnBrowseAttach"
        Me.btnBrowseAttach.Size = New System.Drawing.Size(80, 24)
        Me.btnBrowseAttach.TabIndex = 5
        Me.btnBrowseAttach.Text = "参照..."
        '
        'grpAutoImport
        '
        Me.chkSyncDeletions = New System.Windows.Forms.CheckBox()
        Me.lblSyncMode = New System.Windows.Forms.Label()
        Me.cboSyncMode = New System.Windows.Forms.ComboBox()
        Me.lblDiffBuffer = New System.Windows.Forms.Label()
        Me.numDiffBuffer = New System.Windows.Forms.NumericUpDown()
        Me.lblDiffBufferUnit = New System.Windows.Forms.Label()
        CType(Me.numDiffBuffer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAutoImport.Controls.Add(Me.chkAutoImportEnabled)
        Me.grpAutoImport.Controls.Add(Me.rdoInterval)
        Me.grpAutoImport.Controls.Add(Me.numInterval)
        Me.grpAutoImport.Controls.Add(Me.rdoScheduled)
        Me.grpAutoImport.Controls.Add(Me.dtpScheduledTime)
        Me.grpAutoImport.Controls.Add(Me.lblMaxCount)
        Me.grpAutoImport.Controls.Add(Me.numMaxCount)
        Me.grpAutoImport.Controls.Add(Me.lblImportOrder)
        Me.grpAutoImport.Controls.Add(Me.cboImportOrder)
        Me.grpAutoImport.Controls.Add(Me.lblSyncMode)
        Me.grpAutoImport.Controls.Add(Me.cboSyncMode)
        Me.grpAutoImport.Controls.Add(Me.lblDiffBuffer)
        Me.grpAutoImport.Controls.Add(Me.numDiffBuffer)
        Me.grpAutoImport.Controls.Add(Me.lblDiffBufferUnit)
        Me.grpAutoImport.Controls.Add(Me.chkSyncDeletions)
        Me.grpAutoImport.Location = New System.Drawing.Point(8, 110)
        Me.grpAutoImport.Name = "grpAutoImport"
        Me.grpAutoImport.Size = New System.Drawing.Size(488, 236)
        Me.grpAutoImport.TabIndex = 1
        Me.grpAutoImport.TabStop = False
        Me.grpAutoImport.Text = "自動取り込み"
        '
        'chkAutoImportEnabled
        '
        Me.chkAutoImportEnabled.AutoSize = True
        Me.chkAutoImportEnabled.Location = New System.Drawing.Point(8, 22)
        Me.chkAutoImportEnabled.Name = "chkAutoImportEnabled"
        Me.chkAutoImportEnabled.Size = New System.Drawing.Size(216, 17)
        Me.chkAutoImportEnabled.TabIndex = 0
        Me.chkAutoImportEnabled.Text = "起動時に自動取り込みを開始する"
        '
        'rdoInterval
        '
        Me.rdoInterval.AutoSize = True
        Me.rdoInterval.Checked = True
        Me.rdoInterval.Location = New System.Drawing.Point(24, 48)
        Me.rdoInterval.Name = "rdoInterval"
        Me.rdoInterval.Size = New System.Drawing.Size(130, 17)
        Me.rdoInterval.TabIndex = 1
        Me.rdoInterval.TabStop = True
        Me.rdoInterval.Text = "取り込み間隔（分）:"
        '
        'numInterval
        '
        Me.numInterval.Location = New System.Drawing.Point(180, 47)
        Me.numInterval.Maximum = New Decimal(New Integer() {1440, 0, 0, 0})
        Me.numInterval.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numInterval.Name = "numInterval"
        Me.numInterval.Size = New System.Drawing.Size(60, 22)
        Me.numInterval.TabIndex = 2
        Me.numInterval.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'rdoScheduled
        '
        Me.rdoScheduled.AutoSize = True
        Me.rdoScheduled.Location = New System.Drawing.Point(260, 48)
        Me.rdoScheduled.Name = "rdoScheduled"
        Me.rdoScheduled.Size = New System.Drawing.Size(93, 17)
        Me.rdoScheduled.TabIndex = 3
        Me.rdoScheduled.Text = "定時取り込み:"
        '
        'dtpScheduledTime
        '
        Me.dtpScheduledTime.CustomFormat = "HH:mm"
        Me.dtpScheduledTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpScheduledTime.Location = New System.Drawing.Point(380, 46)
        Me.dtpScheduledTime.Name = "dtpScheduledTime"
        Me.dtpScheduledTime.ShowUpDown = True
        Me.dtpScheduledTime.Size = New System.Drawing.Size(60, 22)
        Me.dtpScheduledTime.TabIndex = 4
        Me.dtpScheduledTime.Value = New System.DateTime(2026, 1, 1, 20, 0, 0, 0)
        '
        'lblMaxCount
        '
        Me.lblMaxCount.AutoSize = True
        Me.lblMaxCount.Location = New System.Drawing.Point(8, 76)
        Me.lblMaxCount.Name = "lblMaxCount"
        Me.lblMaxCount.Size = New System.Drawing.Size(144, 13)
        Me.lblMaxCount.Text = "1 回の最大取り込み件数:"
        '
        'numMaxCount
        '
        Me.numMaxCount.Location = New System.Drawing.Point(180, 73)
        Me.numMaxCount.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.numMaxCount.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numMaxCount.Name = "numMaxCount"
        Me.numMaxCount.Size = New System.Drawing.Size(60, 22)
        Me.numMaxCount.TabIndex = 5
        Me.numMaxCount.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'lblImportOrder
        '
        Me.lblImportOrder.AutoSize = True
        Me.lblImportOrder.Location = New System.Drawing.Point(8, 102)
        Me.lblImportOrder.Name = "lblImportOrder"
        Me.lblImportOrder.Size = New System.Drawing.Size(93, 13)
        Me.lblImportOrder.Text = "取り込み順序:"
        '
        'cboImportOrder
        '
        Me.cboImportOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboImportOrder.FormattingEnabled = True
        Me.cboImportOrder.Items.AddRange(New Object() {"古い順（推奨）", "新しい順"})
        Me.cboImportOrder.Location = New System.Drawing.Point(180, 99)
        Me.cboImportOrder.Name = "cboImportOrder"
        Me.cboImportOrder.Size = New System.Drawing.Size(140, 21)
        Me.cboImportOrder.TabIndex = 6
        '
        'lblSyncMode
        '
        Me.lblSyncMode.AutoSize = True
        Me.lblSyncMode.Location = New System.Drawing.Point(8, 128)
        Me.lblSyncMode.Name = "lblSyncMode"
        Me.lblSyncMode.Size = New System.Drawing.Size(93, 13)
        Me.lblSyncMode.Text = "同期モード:"
        '
        'cboSyncMode
        '
        Me.cboSyncMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSyncMode.FormattingEnabled = True
        Me.cboSyncMode.Items.AddRange(New Object() {"フルスキャン", "差分スキャン（推奨）"})
        Me.cboSyncMode.Location = New System.Drawing.Point(180, 125)
        Me.cboSyncMode.Name = "cboSyncMode"
        Me.cboSyncMode.Size = New System.Drawing.Size(160, 21)
        Me.cboSyncMode.TabIndex = 7
        '
        'lblDiffBuffer
        '
        Me.lblDiffBuffer.AutoSize = True
        Me.lblDiffBuffer.Location = New System.Drawing.Point(8, 154)
        Me.lblDiffBuffer.Name = "lblDiffBuffer"
        Me.lblDiffBuffer.Size = New System.Drawing.Size(144, 13)
        Me.lblDiffBuffer.Text = "差分スキャンのバッファ:"
        '
        'numDiffBuffer
        '
        Me.numDiffBuffer.Location = New System.Drawing.Point(180, 151)
        Me.numDiffBuffer.Maximum = New Decimal(New Integer() {720, 0, 0, 0})
        Me.numDiffBuffer.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numDiffBuffer.Name = "numDiffBuffer"
        Me.numDiffBuffer.Size = New System.Drawing.Size(60, 22)
        Me.numDiffBuffer.TabIndex = 8
        Me.numDiffBuffer.Value = New Decimal(New Integer() {24, 0, 0, 0})
        '
        'lblDiffBufferUnit
        '
        Me.lblDiffBufferUnit.AutoSize = True
        Me.lblDiffBufferUnit.Location = New System.Drawing.Point(244, 154)
        Me.lblDiffBufferUnit.Name = "lblDiffBufferUnit"
        Me.lblDiffBufferUnit.Size = New System.Drawing.Size(30, 13)
        Me.lblDiffBufferUnit.Text = "時間"
        '
        'chkSyncDeletions
        '
        Me.chkSyncDeletions.AutoSize = True
        Me.chkSyncDeletions.Location = New System.Drawing.Point(8, 180)
        Me.chkSyncDeletions.Name = "chkSyncDeletions"
        Me.chkSyncDeletions.Size = New System.Drawing.Size(380, 17)
        Me.chkSyncDeletions.TabIndex = 9
        Me.chkSyncDeletions.Text = "Outlook 側で削除されたメールをデータベースからも削除する"
        '
        'grpFolders
        '
        Me.grpFolders.Controls.Add(Me.lstFolders)
        Me.grpFolders.Controls.Add(Me.btnSelectFolders)
        Me.grpFolders.Location = New System.Drawing.Point(8, 354)
        Me.grpFolders.Name = "grpFolders"
        Me.grpFolders.Size = New System.Drawing.Size(488, 130)
        Me.grpFolders.TabIndex = 2
        Me.grpFolders.TabStop = False
        Me.grpFolders.Text = "対象フォルダ"
        '
        'lstFolders
        '
        Me.lstFolders.FormattingEnabled = True
        Me.lstFolders.Location = New System.Drawing.Point(8, 22)
        Me.lstFolders.Name = "lstFolders"
        Me.lstFolders.Size = New System.Drawing.Size(376, 98)
        Me.lstFolders.TabIndex = 0
        '
        'btnSelectFolders
        '
        Me.btnSelectFolders.Location = New System.Drawing.Point(392, 22)
        Me.btnSelectFolders.Name = "btnSelectFolders"
        Me.btnSelectFolders.Size = New System.Drawing.Size(88, 26)
        Me.btnSelectFolders.TabIndex = 1
        Me.btnSelectFolders.Text = "選択..."
        '
        'grpDisplay
        '
        Me.grpDisplay.Controls.Add(Me.chkDefaultHtml)
        Me.grpDisplay.Controls.Add(Me.chkSortAscending)
        Me.grpDisplay.Controls.Add(Me.chkShowImportResult)
        Me.grpDisplay.Location = New System.Drawing.Point(8, 492)
        Me.grpDisplay.Name = "grpDisplay"
        Me.grpDisplay.Size = New System.Drawing.Size(488, 94)
        Me.grpDisplay.TabIndex = 3
        Me.grpDisplay.TabStop = False
        Me.grpDisplay.Text = "表示設定"
        '
        'chkDefaultHtml
        '
        Me.chkDefaultHtml.AutoSize = True
        Me.chkDefaultHtml.Location = New System.Drawing.Point(8, 22)
        Me.chkDefaultHtml.Name = "chkDefaultHtml"
        Me.chkDefaultHtml.Size = New System.Drawing.Size(208, 17)
        Me.chkDefaultHtml.TabIndex = 0
        Me.chkDefaultHtml.Text = "デフォルトで HTML 表示する"
        '
        'chkSortAscending
        '
        Me.chkSortAscending.AutoSize = True
        Me.chkSortAscending.Location = New System.Drawing.Point(8, 46)
        Me.chkSortAscending.Name = "chkSortAscending"
        Me.chkSortAscending.Size = New System.Drawing.Size(264, 17)
        Me.chkSortAscending.TabIndex = 1
        Me.chkSortAscending.Text = "会話ビューを古い順（昇順）で表示する"
        '
        'chkShowImportResult
        '
        Me.chkShowImportResult.AutoSize = True
        Me.chkShowImportResult.Location = New System.Drawing.Point(8, 70)
        Me.chkShowImportResult.Name = "chkShowImportResult"
        Me.chkShowImportResult.Size = New System.Drawing.Size(296, 17)
        Me.chkShowImportResult.TabIndex = 2
        Me.chkShowImportResult.Text = "取り込み完了時に結果ダイアログを表示する"
        '
        'grpTray
        '
        Me.chkStartWithWindows = New System.Windows.Forms.CheckBox()
        Me.grpTray.Controls.Add(Me.chkStartWithWindows)
        Me.grpTray.Controls.Add(Me.chkMinimizeToTray)
        Me.grpTray.Controls.Add(Me.chkCloseToTray)
        Me.grpTray.Controls.Add(Me.chkShowBalloonOnImport)
        Me.grpTray.Location = New System.Drawing.Point(8, 594)
        Me.grpTray.Name = "grpTray"
        Me.grpTray.Size = New System.Drawing.Size(488, 118)
        Me.grpTray.TabIndex = 6
        Me.grpTray.TabStop = False
        Me.grpTray.Text = "タスクトレイ"
        '
        'chkStartWithWindows
        '
        Me.chkStartWithWindows.AutoSize = True
        Me.chkStartWithWindows.Location = New System.Drawing.Point(8, 22)
        Me.chkStartWithWindows.Name = "chkStartWithWindows"
        Me.chkStartWithWindows.Size = New System.Drawing.Size(296, 17)
        Me.chkStartWithWindows.TabIndex = 0
        Me.chkStartWithWindows.Text = "Windows 起動時に自動起動する（トレイに常駐）"
        '
        'chkMinimizeToTray
        '
        Me.chkMinimizeToTray.AutoSize = True
        Me.chkMinimizeToTray.Location = New System.Drawing.Point(8, 46)
        Me.chkMinimizeToTray.Name = "chkMinimizeToTray"
        Me.chkMinimizeToTray.Size = New System.Drawing.Size(296, 17)
        Me.chkMinimizeToTray.TabIndex = 1
        Me.chkMinimizeToTray.Text = "最小化時にタスクトレイに格納する"
        '
        'chkCloseToTray
        '
        Me.chkCloseToTray.AutoSize = True
        Me.chkCloseToTray.Location = New System.Drawing.Point(8, 70)
        Me.chkCloseToTray.Name = "chkCloseToTray"
        Me.chkCloseToTray.Size = New System.Drawing.Size(296, 17)
        Me.chkCloseToTray.TabIndex = 2
        Me.chkCloseToTray.Text = "閉じるボタンでタスクトレイに格納する"
        '
        'chkShowBalloonOnImport
        '
        Me.chkShowBalloonOnImport.AutoSize = True
        Me.chkShowBalloonOnImport.Location = New System.Drawing.Point(8, 94)
        Me.chkShowBalloonOnImport.Name = "chkShowBalloonOnImport"
        Me.chkShowBalloonOnImport.Size = New System.Drawing.Size(296, 17)
        Me.chkShowBalloonOnImport.TabIndex = 3
        Me.chkShowBalloonOnImport.Text = "取り込み完了時にバルーン通知を表示する"
        '
        'grpData
        '
        Me.grpData.Controls.Add(Me.btnResetData)
        Me.grpData.Location = New System.Drawing.Point(8, 720)
        Me.grpData.Name = "grpData"
        Me.grpData.Size = New System.Drawing.Size(488, 56)
        Me.grpData.TabIndex = 5
        Me.grpData.TabStop = False
        Me.grpData.Text = "データ管理"
        '
        'btnResetData
        '
        Me.btnResetData.Location = New System.Drawing.Point(8, 22)
        Me.btnResetData.Name = "btnResetData"
        Me.btnResetData.Size = New System.Drawing.Size(120, 26)
        Me.btnResetData.TabIndex = 0
        Me.btnResetData.Text = "データ初期化..."
        '
        'pnlButtons
        '
        Me.pnlButtons.Controls.Add(Me.btnOk)
        Me.pnlButtons.Controls.Add(Me.btnCancel)
        Me.pnlButtons.Location = New System.Drawing.Point(8, 784)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Size = New System.Drawing.Size(488, 34)
        Me.pnlButtons.TabIndex = 4
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(324, 4)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(72, 26)
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(404, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "キャンセル"
        '
        'SettingsForm
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Font = New System.Drawing.Font("Meiryo UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(512, 828)
        Me.Controls.Add(Me.grpPaths)
        Me.Controls.Add(Me.grpAutoImport)
        Me.Controls.Add(Me.grpFolders)
        Me.Controls.Add(Me.grpDisplay)
        Me.Controls.Add(Me.grpTray)
        Me.Controls.Add(Me.grpData)
        Me.Controls.Add(Me.pnlButtons)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "設定"
        Me.grpPaths.ResumeLayout(False)
        Me.grpPaths.PerformLayout()
        Me.grpAutoImport.ResumeLayout(False)
        Me.grpAutoImport.PerformLayout()
        CType(Me.numInterval, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numMaxCount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numDiffBuffer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFolders.ResumeLayout(False)
        Me.grpDisplay.ResumeLayout(False)
        Me.grpDisplay.PerformLayout()
        Me.grpTray.ResumeLayout(False)
        Me.grpTray.PerformLayout()
        Me.grpData.ResumeLayout(False)
        Me.pnlButtons.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    ' ── コントロール宣言 ──────────────────────────────────────────
    Friend WithEvents grpPaths As System.Windows.Forms.GroupBox
    Friend WithEvents lblDbPath As System.Windows.Forms.Label
    Friend WithEvents txtDbPath As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseDb As System.Windows.Forms.Button
    Friend WithEvents lblAttachDir As System.Windows.Forms.Label
    Friend WithEvents txtAttachDir As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseAttach As System.Windows.Forms.Button
    Friend WithEvents grpAutoImport As System.Windows.Forms.GroupBox
    Friend WithEvents chkAutoImportEnabled As System.Windows.Forms.CheckBox
    Friend WithEvents rdoInterval As System.Windows.Forms.RadioButton
    Friend WithEvents rdoScheduled As System.Windows.Forms.RadioButton
    Friend WithEvents numInterval As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtpScheduledTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblMaxCount As System.Windows.Forms.Label
    Friend WithEvents numMaxCount As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblImportOrder As System.Windows.Forms.Label
    Friend WithEvents cboImportOrder As System.Windows.Forms.ComboBox
    Friend WithEvents lblSyncMode As System.Windows.Forms.Label
    Friend WithEvents cboSyncMode As System.Windows.Forms.ComboBox
    Friend WithEvents lblDiffBuffer As System.Windows.Forms.Label
    Friend WithEvents numDiffBuffer As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblDiffBufferUnit As System.Windows.Forms.Label
    Friend WithEvents grpFolders As System.Windows.Forms.GroupBox
    Friend WithEvents lstFolders As System.Windows.Forms.ListBox
    Friend WithEvents btnSelectFolders As System.Windows.Forms.Button
    Friend WithEvents grpDisplay As System.Windows.Forms.GroupBox
    Friend WithEvents chkDefaultHtml As System.Windows.Forms.CheckBox
    Friend WithEvents chkSortAscending As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowImportResult As System.Windows.Forms.CheckBox
    Friend WithEvents chkSyncDeletions As System.Windows.Forms.CheckBox
    Friend WithEvents grpTray As System.Windows.Forms.GroupBox
    Friend WithEvents chkMinimizeToTray As System.Windows.Forms.CheckBox
    Friend WithEvents chkCloseToTray As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowBalloonOnImport As System.Windows.Forms.CheckBox
    Friend WithEvents chkStartWithWindows As System.Windows.Forms.CheckBox
    Friend WithEvents grpData As System.Windows.Forms.GroupBox
    Friend WithEvents btnResetData As System.Windows.Forms.Button
    Friend WithEvents pnlButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button

End Class
