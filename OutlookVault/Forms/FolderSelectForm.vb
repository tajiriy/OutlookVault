Option Explicit On
Option Strict On
Option Infer Off

''' <summary>
''' Outlook のフォルダ一覧を表示し、対象フォルダをチェックボックスで選択するダイアログ。
''' </summary>
Partial Public Class FolderSelectForm

    Private ReadOnly _allFolders As List(Of String)
    Private ReadOnly _selectedFolders As HashSet(Of String)

    ''' <summary>ダイアログ確定後に選択されたフォルダ名のリストを返す。</summary>
    Public ReadOnly Property SelectedFolderNames As List(Of String)
        Get
            Dim result As New List(Of String)()
            For i As Integer = 0 To clbFolders.Items.Count - 1
                If clbFolders.GetItemChecked(i) Then
                    result.Add(CStr(clbFolders.Items(i)))
                End If
            Next
            ' フィルタで非表示のフォルダのうちチェック済みだったものも含める
            For Each folder As String In _selectedFolders
                If Not result.Contains(folder) Then
                    ' フィルタで非表示になっているが選択済みだったフォルダ
                    ' 現在リストに表示されていないものを追加
                    Dim isInList As Boolean = False
                    For j As Integer = 0 To clbFolders.Items.Count - 1
                        If CStr(clbFolders.Items(j)) = folder Then
                            isInList = True
                            Exit For
                        End If
                    Next
                    If Not isInList Then
                        result.Add(folder)
                    End If
                End If
            Next
            Return result
        End Get
    End Property

    ''' <summary>
    ''' コンストラクタ。
    ''' </summary>
    ''' <param name="availableFolders">Outlook から取得したフォルダ名一覧。</param>
    ''' <param name="currentSelection">現在選択中のフォルダ名一覧。</param>
    Public Sub New(availableFolders As List(Of String), currentSelection As List(Of String))
        _allFolders = availableFolders
        _selectedFolders = New HashSet(Of String)(If(currentSelection, New List(Of String)()),
                                                   StringComparer.OrdinalIgnoreCase)
        InitializeComponent()
        Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
        If appIcon IsNot Nothing Then Me.Icon = appIcon
        PopulateFolderList(String.Empty)
    End Sub

    ''' <summary>フィルタ文字列に基づいてフォルダ一覧を更新する。</summary>
    Private Sub PopulateFolderList(filter As String)
        ' 現在のチェック状態を保存
        SaveCheckState()

        clbFolders.Items.Clear()
        Dim checkedCount As Integer = 0

        For Each folder As String In _allFolders
            If filter.Length > 0 AndAlso folder.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0 Then
                Continue For
            End If
            Dim idx As Integer = clbFolders.Items.Add(CType(folder, Object))
            If _selectedFolders.Contains(folder) Then
                clbFolders.SetItemChecked(idx, True)
                checkedCount += 1
            End If
        Next

        UpdateStatus(checkedCount)
    End Sub

    ''' <summary>CheckedListBox の現在のチェック状態を _selectedFolders に反映する。</summary>
    Private Sub SaveCheckState()
        For i As Integer = 0 To clbFolders.Items.Count - 1
            Dim name As String = CStr(clbFolders.Items(i))
            If clbFolders.GetItemChecked(i) Then
                If Not _selectedFolders.Contains(name) Then
                    _selectedFolders.Add(name)
                End If
            Else
                _selectedFolders.Remove(name)
            End If
        Next
    End Sub

    Private Sub UpdateStatus(checkedCount As Integer)
        lblStatus.Text = checkedCount.ToString() & " 件選択中 / " & _allFolders.Count.ToString() & " 件"
    End Sub

    ' ── イベントハンドラ ──

    Private Sub txtFilter_TextChanged(sender As Object, e As System.EventArgs) Handles txtFilter.TextChanged
        PopulateFolderList(txtFilter.Text.Trim())
    End Sub

    Private Sub btnOk_Click(sender As Object, e As System.EventArgs) Handles btnOk.Click
        SaveCheckState()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
