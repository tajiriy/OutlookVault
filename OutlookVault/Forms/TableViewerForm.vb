Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data
Imports System.Data.SQLite
Imports System.Text
Imports System.Windows.Forms

Namespace Forms

    ''' <summary>
    ''' データベーステーブルの内容を DataGridView で表示するフォーム。
    ''' 列ヘッダクリックでソート、テキストボックスで行フィルタ、プルダウンでテーブル切替が可能。
    ''' </summary>
    Partial Public Class TableViewerForm

        Private Shared ReadOnly TableNames() As String = {"emails", "attachments", "deleted_message_ids", "exchange_address_cache", "error_message_ids", "folder_sync_state"}
        Private Const MaxColumnWidth As Integer = 200

        Private ReadOnly _dbManager As Data.DatabaseManager

        Private Const FilterDelayMs As Integer = 800

        Private _dataTable As DataTable
        Private _bindingSource As BindingSource

        Public Sub New(dbManager As Data.DatabaseManager, tableName As String)
            _dbManager = dbManager
            Me.DoubleBuffered = True
            _bindingSource = New BindingSource()
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon
            cboTable.Items.AddRange(DirectCast(TableNames, Object()))
            dgv.DataSource = _bindingSource
            AddHandler txtFilter.TextChanged, AddressOf TxtFilter_TextChanged
            AddHandler _filterTimer.Tick, AddressOf FilterTimer_Tick
            ' 初期テーブルを選択（イベント経由で LoadData が走る）
            cboTable.SelectedItem = tableName
        End Sub

        ''' <summary>
        ''' スクロール・リサイズ時のゴースト表示を防止する DataGridView サブクラス。
        ''' DoubleBuffered を有効化し、スクロール/リサイズ時に全面再描画を強制する。
        ''' </summary>
        Private Class BufferedDataGridView
            Inherits DataGridView

            Public Sub New()
                Me.DoubleBuffered = True
                SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
            End Sub

            Protected Overrides Sub OnResize(e As EventArgs)
                MyBase.OnResize(e)
                Me.Invalidate()
            End Sub

        End Class

        Private Sub cboTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTable.SelectedIndexChanged
            If cboTable.SelectedItem Is Nothing Then Return
            Dim selectedTable As String = CStr(cboTable.SelectedItem)
            Me.Text = "テーブルビューア - " & selectedTable
            txtFilter.Text = ""
            LoadData(selectedTable)
        End Sub

        Private Sub LoadData(tableName As String)
            dgv.SuspendLayout()
            Try
                If _dataTable IsNot Nothing Then
                    _bindingSource.DataSource = Nothing
                    _dataTable.Dispose()
                End If
                _dataTable = _dbManager.GetTableData(tableName)
                _bindingSource.DataSource = _dataTable
                ApplyMaxColumnWidth()
                UpdateRowCount()
            Finally
                dgv.ResumeLayout()
            End Try
        End Sub

        Private Sub ApplyMaxColumnWidth()
            ' 表示中のセルに基づいて1回だけ自動サイズ計算し、上限を適用
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
            For Each col As DataGridViewColumn In dgv.Columns
                If col.Width > MaxColumnWidth Then
                    col.Width = MaxColumnWidth
                End If
                ' ユーザーによるドラッグリサイズを確実に有効化
                col.Resizable = DataGridViewTriState.True
            Next
        End Sub

        Private Sub TxtFilter_TextChanged(sender As Object, e As EventArgs)
            ' キー入力のたびにタイマーをリセットして再スタート
            _filterTimer.Stop()
            _filterTimer.Start()
        End Sub

        Private Sub FilterTimer_Tick(sender As Object, e As EventArgs)
            _filterTimer.Stop()
            ApplyFilter()
        End Sub

        Private Sub ApplyFilter()
            Dim filterText As String = txtFilter.Text.Trim()
            If String.IsNullOrEmpty(filterText) Then
                _dataTable.DefaultView.RowFilter = ""
                UpdateRowCount()
                Return
            End If

            Try
                _dataTable.DefaultView.RowFilter = Filters.FilterParser.Parse(filterText, _dataTable.Columns)
            Catch ex As Exception
                _dataTable.DefaultView.RowFilter = ""
            End Try

            UpdateRowCount()
        End Sub

        Private Sub UpdateRowCount()
            Dim total As Integer = _dataTable.Rows.Count
            Dim filtered As Integer = _dataTable.DefaultView.Count
            If total = filtered Then
                lblRowCount.Text = String.Format("{0:N0} 件", total)
            Else
                lblRowCount.Text = String.Format("{0:N0} / {1:N0} 件", filtered, total)
            End If
            ' 右上に配置
            lblRowCount.Location = New Drawing.Point(
                pnlTop.ClientSize.Width - lblRowCount.Width - 10, 10)
        End Sub

        Private Sub tsmiCopy_Click(sender As Object, e As EventArgs) Handles tsmiCopy.Click
            CopySelectedCells()
        End Sub

        Private Sub dgv_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv.KeyDown
            If e.Control AndAlso e.KeyCode = Keys.C Then
                CopySelectedCells()
                e.Handled = True
            End If
        End Sub

        Private Sub CopySelectedCells()
            If dgv.SelectedCells.Count = 0 Then Return
            Dim text As String = BuildCopyText(dgv.SelectedCells)
            If text.Length > 0 Then
                Clipboard.SetText(text)
            End If
        End Sub

        ''' <summary>
        ''' 選択セルからコピー用テキストを生成する。
        ''' 矩形選択（連続した行×列）の場合は横タブ・縦改行。
        ''' とびとびの選択の場合はタブ区切りのフラット出力。
        ''' </summary>
        Public Shared Function BuildCopyText(cells As DataGridViewSelectedCellCollection) As String
            If cells.Count = 0 Then Return ""
            If cells.Count = 1 Then
                Return If(cells(0).Value IsNot Nothing, cells(0).Value.ToString(), "")
            End If

            ' 選択セルの行・列インデックスを収集
            Dim rows As New SortedSet(Of Integer)()
            Dim cols As New SortedSet(Of Integer)()
            Dim cellMap As New Dictionary(Of Long, DataGridViewCell)()
            For Each cell As DataGridViewCell In cells
                rows.Add(cell.RowIndex)
                cols.Add(cell.ColumnIndex)
                Dim key As Long = (CLng(cell.RowIndex) << 20) Or CLng(cell.ColumnIndex)
                cellMap(key) = cell
            Next

            ' 矩形判定: 行数×列数 = 選択セル数
            Dim isRectangular As Boolean = (rows.Count * cols.Count = cells.Count)

            If isRectangular AndAlso rows.Count > 1 Then
                ' 矩形選択: 横タブ・縦改行
                Dim sb As New StringBuilder()
                Dim firstRow As Boolean = True
                For Each r As Integer In rows
                    If Not firstRow Then sb.Append(vbCrLf)
                    firstRow = False
                    Dim firstCol As Boolean = True
                    For Each c As Integer In cols
                        If Not firstCol Then sb.Append(vbTab)
                        firstCol = False
                        Dim key As Long = (CLng(r) << 20) Or CLng(c)
                        Dim cell As DataGridViewCell = Nothing
                        If cellMap.TryGetValue(key, cell) Then
                            sb.Append(If(cell.Value IsNot Nothing, cell.Value.ToString(), ""))
                        End If
                    Next
                Next
                Return sb.ToString()
            Else
                ' とびとび選択 or 単一行: タブ区切りフラット
                ' 行→列の順にソートして出力
                Dim sb As New StringBuilder()
                Dim first As Boolean = True
                For Each r As Integer In rows
                    For Each c As Integer In cols
                        Dim key As Long = (CLng(r) << 20) Or CLng(c)
                        Dim cell As DataGridViewCell = Nothing
                        If cellMap.TryGetValue(key, cell) Then
                            If Not first Then sb.Append(vbTab)
                            first = False
                            sb.Append(If(cell.Value IsNot Nothing, cell.Value.ToString(), ""))
                        End If
                    Next
                Next
                Return sb.ToString()
            End If
        End Function

        Private Sub dgv_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv.ColumnHeaderMouseClick
            Dim colName As String = dgv.Columns(e.ColumnIndex).DataPropertyName
            Dim currentSort As String = If(_dataTable.DefaultView.Sort, "")

            If currentSort.StartsWith(String.Format("[{0}]", colName), StringComparison.OrdinalIgnoreCase) AndAlso
               Not currentSort.EndsWith("DESC", StringComparison.OrdinalIgnoreCase) Then
                _dataTable.DefaultView.Sort = String.Format("[{0}] DESC", colName)
            Else
                _dataTable.DefaultView.Sort = String.Format("[{0}] ASC", colName)
            End If
        End Sub

        Private Sub TableViewerForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
            If _bindingSource IsNot Nothing Then _bindingSource.Dispose()
            If _dataTable IsNot Nothing Then _dataTable.Dispose()
        End Sub

    End Class

End Namespace
