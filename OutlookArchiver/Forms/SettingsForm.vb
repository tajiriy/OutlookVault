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
        numInterval.Value = CDec(Math.Min(CInt(numInterval.Maximum), Math.Max(CInt(numInterval.Minimum), _settings.AutoImportIntervalMinutes)))
        numMaxCount.Value = CDec(Math.Min(CInt(numMaxCount.Maximum), Math.Max(CInt(numMaxCount.Minimum), _settings.MaxImportCount)))

        lstFolders.Items.Clear()
        For Each folder As String In _settings.TargetFolders
            If Not String.IsNullOrEmpty(folder) Then
                lstFolders.Items.Add(CType(folder, Object))
            End If
        Next

        cboImportOrder.SelectedIndex = If(_settings.ImportOldestFirst, 0, 1)

        chkDefaultHtml.Checked = _settings.DefaultHtmlView
        chkSortAscending.Checked = _settings.ConversationSortAscending
        chkShowImportResult.Checked = _settings.ShowImportResult
    End Sub

    ''' <summary>各コントロールの値を設定に保存する。</summary>
    Private Sub SaveSettings()
        _settings.DbFilePath = txtDbPath.Text.Trim()
        _settings.AttachmentDirectory = txtAttachDir.Text.Trim()
        _settings.AutoImportEnabled = chkAutoImportEnabled.Checked
        _settings.AutoImportIntervalMinutes = CInt(numInterval.Value)
        _settings.MaxImportCount = CInt(numMaxCount.Value)

        Dim folders As New List(Of String)()
        For Each item As Object In lstFolders.Items
            Dim s As String = TryCast(item, String)
            If s IsNot Nothing AndAlso s.Length > 0 Then folders.Add(s)
        Next
        _settings.TargetFolders = folders

        _settings.ImportOldestFirst = (cboImportOrder.SelectedIndex = 0)

        _settings.DefaultHtmlView = chkDefaultHtml.Checked
        _settings.ConversationSortAscending = chkSortAscending.Checked
        _settings.ShowImportResult = chkShowImportResult.Checked
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

    Private Sub btnAddFolder_Click(sender As Object, e As System.EventArgs) Handles btnAddFolder.Click
        Dim name As String = Microsoft.VisualBasic.Interaction.InputBox(
            "追加する Outlook フォルダ名を入力してください:", "フォルダ追加", String.Empty)
        name = name.Trim()
        If name.Length > 0 AndAlso Not lstFolders.Items.Contains(CType(name, Object)) Then
            lstFolders.Items.Add(CType(name, Object))
        End If
    End Sub

    Private Sub btnRemoveFolder_Click(sender As Object, e As System.EventArgs) Handles btnRemoveFolder.Click
        If lstFolders.SelectedIndex >= 0 Then
            lstFolders.Items.RemoveAt(lstFolders.SelectedIndex)
        End If
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
