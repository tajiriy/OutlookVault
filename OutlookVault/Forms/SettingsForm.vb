Option Explicit On
Option Strict On
Option Infer Off

''' <summary>
''' アプリケーション設定を編集するダイアログフォーム。
''' OK ボタンで設定を保存し、キャンセルで破棄する。
''' </summary>
Public Class SettingsForm

    Private ReadOnly _settings As Config.AppSettings

    ''' <summary>データ初期化が正常完了した場合 True。MainForm 側で再初期化の判断に使用する。</summary>
    Public Property DataWasReset As Boolean = False

    ' ════════════════════════════════════════════════════════════
    '  初期化
    ' ════════════════════════════════════════════════════════════

    Public Sub New()
        InitializeComponent()
        _settings = Config.AppSettings.Instance
        LoadSettings()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  設定の読み込み・保存
    ' ════════════════════════════════════════════════════════════

    ''' <summary>現在の設定値を各コントロールに反映する。</summary>
    Private Sub LoadSettings()
        txtDbPath.Text = _settings.DbFilePath
        txtAttachDir.Text = _settings.AttachmentDirectory
        chkAutoImportEnabled.Checked = _settings.AutoImportEnabled
        If _settings.AutoImportMode = 1 Then
            rdoScheduled.Checked = True
        Else
            rdoInterval.Checked = True
        End If
        numInterval.Value = CDec(Math.Min(CInt(numInterval.Maximum), Math.Max(CInt(numInterval.Minimum), _settings.AutoImportIntervalMinutes)))

        ' 定時取り込み時刻を DateTimePicker に設定
        Dim scheduledTime As DateTime
        If DateTime.TryParseExact(_settings.ScheduledImportTime, "HH:mm",
                                   System.Globalization.CultureInfo.InvariantCulture,
                                   System.Globalization.DateTimeStyles.None, scheduledTime) Then
            dtpScheduledTime.Value = scheduledTime
        End If

        UpdateImportModeControls()
        numMaxCount.Value = CDec(Math.Min(CInt(numMaxCount.Maximum), Math.Max(CInt(numMaxCount.Minimum), _settings.MaxImportCount)))

        lstFolders.Items.Clear()
        For Each folder As String In _settings.TargetFolders
            If Not String.IsNullOrEmpty(folder) Then
                lstFolders.Items.Add(CType(folder, Object))
            End If
        Next

        cboImportOrder.SelectedIndex = If(_settings.ImportOldestFirst, 0, 1)
        cboSyncMode.SelectedIndex = _settings.SyncMode
        numDiffBuffer.Value = CDec(Math.Min(CInt(numDiffBuffer.Maximum), Math.Max(CInt(numDiffBuffer.Minimum), _settings.DiffSyncBufferHours)))
        UpdateSyncModeControls()
        chkSyncDeletions.Checked = _settings.SyncDeletionsEnabled

        chkDefaultHtml.Checked = _settings.DefaultHtmlView
        chkSortAscending.Checked = _settings.ConversationSortAscending
        chkShowImportResult.Checked = _settings.ShowImportResult
        chkShowImportErrorDialog.Checked = _settings.ShowImportErrorDialog

        chkStartWithWindows.Checked = _settings.StartWithWindows
        chkMinimizeToTray.Checked = _settings.MinimizeToTray
        chkCloseToTray.Checked = _settings.CloseToTray
        chkShowBalloonOnImport.Checked = _settings.ShowBalloonOnImport
    End Sub

    ''' <summary>各コントロールの値を設定に保存する。</summary>
    Private Sub SaveSettings()
        _settings.DbFilePath = txtDbPath.Text.Trim()
        _settings.AttachmentDirectory = txtAttachDir.Text.Trim()
        _settings.AutoImportEnabled = chkAutoImportEnabled.Checked
        _settings.AutoImportMode = If(rdoScheduled.Checked, 1, 0)
        _settings.AutoImportIntervalMinutes = CInt(numInterval.Value)
        _settings.ScheduledImportTime = dtpScheduledTime.Value.ToString("HH:mm")
        _settings.MaxImportCount = CInt(numMaxCount.Value)

        Dim folders As New List(Of String)()
        For Each item As Object In lstFolders.Items
            Dim s As String = TryCast(item, String)
            If s IsNot Nothing AndAlso s.Length > 0 Then folders.Add(s)
        Next
        _settings.TargetFolders = folders

        _settings.ImportOldestFirst = (cboImportOrder.SelectedIndex = 0)
        _settings.SyncMode = cboSyncMode.SelectedIndex
        _settings.DiffSyncBufferHours = CInt(numDiffBuffer.Value)
        _settings.SyncDeletionsEnabled = chkSyncDeletions.Checked

        _settings.DefaultHtmlView = chkDefaultHtml.Checked
        _settings.ConversationSortAscending = chkSortAscending.Checked
        _settings.ShowImportResult = chkShowImportResult.Checked
        _settings.ShowImportErrorDialog = chkShowImportErrorDialog.Checked

        _settings.StartWithWindows = chkStartWithWindows.Checked
        _settings.MinimizeToTray = chkMinimizeToTray.Checked
        _settings.CloseToTray = chkCloseToTray.Checked
        _settings.ShowBalloonOnImport = chkShowBalloonOnImport.Checked
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ 自動取り込みモード
    ' ════════════════════════════════════════════════════════════

    ''' <summary>ラジオボタンの選択に応じてコントロールの有効/無効を切り替える。</summary>
    Private Sub UpdateImportModeControls()
        numInterval.Enabled = rdoInterval.Checked
        dtpScheduledTime.Enabled = rdoScheduled.Checked
        ' 間隔モードのみ最大件数を設定可能（定時は完全同期のため無制限）
        numMaxCount.Enabled = rdoInterval.Checked
    End Sub

    Private Sub rdoInterval_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoInterval.CheckedChanged
        UpdateImportModeControls()
    End Sub

    Private Sub rdoScheduled_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoScheduled.CheckedChanged
        UpdateImportModeControls()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ 同期モード
    ' ════════════════════════════════════════════════════════════

    ''' <summary>同期モードの選択に応じてバッファ設定の有効/無効を切り替える。</summary>
    Private Sub UpdateSyncModeControls()
        Dim isDiff As Boolean = (cboSyncMode.SelectedIndex = 1)
        numDiffBuffer.Enabled = isDiff
        lblDiffBuffer.Enabled = isDiff
        lblDiffBufferUnit.Enabled = isDiff
    End Sub

    Private Sub cboSyncMode_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cboSyncMode.SelectedIndexChanged
        UpdateSyncModeControls()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ 保存先
    ' ════════════════════════════════════════════════════════════

    Private Sub btnBrowseDb_Click(sender As Object, e As System.EventArgs) Handles btnBrowseDb.Click
        Using dlg As New System.Windows.Forms.SaveFileDialog()
            dlg.Title = "DB ファイルの保存先を選択"
            dlg.Filter = "SQLite データベース (*.db)|*.db|すべてのファイル (*.*)|*.*"
            dlg.FileName = txtDbPath.Text
            If dlg.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                txtDbPath.Text = dlg.FileName
            End If
        End Using
    End Sub

    Private Sub btnBrowseAttach_Click(sender As Object, e As System.EventArgs) Handles btnBrowseAttach.Click
        Using dlg As New System.Windows.Forms.FolderBrowserDialog()
            dlg.Description = "添付ファイルの保存先フォルダを選択してください"
            If System.IO.Directory.Exists(txtAttachDir.Text) Then
                dlg.SelectedPath = txtAttachDir.Text
            End If
            If dlg.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                txtAttachDir.Text = dlg.SelectedPath
            End If
        End Using
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ 対象フォルダ
    ' ════════════════════════════════════════════════════════════

    Private Sub btnSelectFolders_Click(sender As Object, e As System.EventArgs) Handles btnSelectFolders.Click
        ' Outlook に接続してフォルダ一覧を取得
        Dim outlookSvc As Services.OutlookService = Nothing
        Try
            outlookSvc = Services.OutlookService.TryConnect()
            If outlookSvc Is Nothing Then
                System.Windows.Forms.MessageBox.Show(
                    "Outlook に接続できませんでした。Outlook が起動していることを確認してください。",
                    "接続エラー",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            Dim availableFolders As List(Of String) = outlookSvc.GetAvailableFolderNames()

            ' 現在の選択済みフォルダをリストから取得
            Dim currentSelection As New List(Of String)()
            For Each item As Object In lstFolders.Items
                Dim s As String = TryCast(item, String)
                If s IsNot Nothing AndAlso s.Length > 0 Then currentSelection.Add(s)
            Next

            ' フォルダ選択ダイアログを表示
            Using dlg As New FolderSelectForm(availableFolders, currentSelection)
                If dlg.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    lstFolders.Items.Clear()
                    For Each folder As String In dlg.SelectedFolderNames
                        lstFolders.Items.Add(CType(folder, Object))
                    Next
                End If
            End Using
        Finally
            If outlookSvc IsNot Nothing Then outlookSvc.Dispose()
        End Try
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ データ初期化
    ' ════════════════════════════════════════════════════════════

    Private Sub btnResetData_Click(sender As Object, e As System.EventArgs) Handles btnResetData.Click
        Dim answer As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(
            "すべての取り込み済みメールデータと添付ファイルを削除します。" & vbCrLf &
            "この操作は元に戻せません。" & vbCrLf & vbCrLf &
            "本当に初期化しますか？",
            "データ初期化の確認",
            System.Windows.Forms.MessageBoxButtons.YesNo,
            System.Windows.Forms.MessageBoxIcon.Warning,
            System.Windows.Forms.MessageBoxDefaultButton.Button2)
        If answer <> System.Windows.Forms.DialogResult.Yes Then Return

        Try
            ' DB ファイルを削除
            Dim dbPath As String = System.IO.Path.GetFullPath(_settings.DbFilePath)
            If System.IO.File.Exists(dbPath) Then
                System.IO.File.Delete(dbPath)
            End If

            ' 添付ファイルディレクトリを削除
            Dim attachDir As String = System.IO.Path.GetFullPath(_settings.AttachmentDirectory)
            If System.IO.Directory.Exists(attachDir) Then
                System.IO.Directory.Delete(attachDir, recursive:=True)
            End If

            ' MainForm 側で再初期化を行うためフラグをセットしてダイアログを閉じる
            DataWasReset = True
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(
                "初期化中にエラーが発生しました:" & vbCrLf & ex.Message,
                "エラー",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ ─ OK / キャンセル
    ' ════════════════════════════════════════════════════════════

    Private Sub btnOk_Click(sender As Object, e As System.EventArgs) Handles btnOk.Click
        SaveSettings()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
