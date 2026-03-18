Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data
Imports System.Data.SQLite
Imports System.Windows.Forms

Namespace Forms

    ''' <summary>
    ''' データベーステーブルの内容を DataGridView で表示するフォーム。
    ''' 列ヘッダクリックでソート、テキストボックスで行フィルタ、プルダウンでテーブル切替が可能。
    ''' </summary>
    Public Class TableViewerForm
        Inherits Form

        Private Shared ReadOnly TableNames() As String = {"emails", "attachments", "deleted_message_ids", "exchange_address_cache"}
        Private Const MaxColumnWidth As Integer = 200

        Private ReadOnly _dbManager As Data.DatabaseManager

        Private Const FilterDelayMs As Integer = 800

        Private WithEvents dgv As DataGridView
        Private WithEvents cboTable As ComboBox
        Private txtFilter As TextBox
        Private lblFilter As Label
        Private lblRowCount As Label
        Private pnlTop As Panel
        Private _dataTable As DataTable
        Private _bindingSource As BindingSource
        Private _filterTimer As Timer

        Public Sub New(dbManager As Data.DatabaseManager, tableName As String)
            _dbManager = dbManager
            Me.DoubleBuffered = True
            InitializeComponents()
            ' 初期テーブルを選択（イベント経由で LoadData が走る）
            cboTable.SelectedItem = tableName
        End Sub

        Private Sub InitializeComponents()
            Me.Text = "テーブルビューア"
            Me.Size = New Drawing.Size(1000, 600)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.MinimumSize = New Drawing.Size(600, 400)

            ' ── 上部パネル（テーブル選択 + フィルタ） ──
            pnlTop = New Panel()
            pnlTop.Dock = DockStyle.Top
            pnlTop.Height = 36
            pnlTop.Padding = New Padding(6, 6, 6, 4)

            Dim lblTable As New Label()
            lblTable.Text = "テーブル:"
            lblTable.AutoSize = True
            lblTable.Location = New Drawing.Point(8, 10)

            cboTable = New ComboBox()
            cboTable.DropDownStyle = ComboBoxStyle.DropDownList
            cboTable.Location = New Drawing.Point(66, 7)
            cboTable.Width = 180
            cboTable.Items.AddRange(DirectCast(TableNames, Object()))

            lblFilter = New Label()
            lblFilter.Text = "フィルタ:"
            lblFilter.AutoSize = True
            lblFilter.Location = New Drawing.Point(260, 10)

            txtFilter = New TextBox()
            txtFilter.Location = New Drawing.Point(312, 7)
            txtFilter.Width = 300
            AddHandler txtFilter.TextChanged, AddressOf TxtFilter_TextChanged

            ' フィルタ遅延用タイマー（キー入力のたびにリセット）
            _filterTimer = New Timer()
            _filterTimer.Interval = FilterDelayMs
            AddHandler _filterTimer.Tick, AddressOf FilterTimer_Tick

            lblRowCount = New Label()
            lblRowCount.AutoSize = True
            lblRowCount.Anchor = CType(AnchorStyles.Top Or AnchorStyles.Right, AnchorStyles)
            lblRowCount.Text = ""

            pnlTop.Controls.Add(lblTable)
            pnlTop.Controls.Add(cboTable)
            pnlTop.Controls.Add(lblFilter)
            pnlTop.Controls.Add(txtFilter)
            pnlTop.Controls.Add(lblRowCount)

            ' ── DataGridView ──
            dgv = New DataGridView()
            dgv.Dock = DockStyle.Fill
            dgv.ReadOnly = True
            dgv.AllowUserToAddRows = False
            dgv.AllowUserToDeleteRows = False
            dgv.AllowUserToOrderColumns = True
            dgv.AllowUserToResizeColumns = True
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
            ' ちらつき防止: DoubleBuffered を有効化
            EnableDoubleBuffering(dgv)

            _bindingSource = New BindingSource()
            dgv.DataSource = _bindingSource

            Me.Controls.Add(dgv)
            Me.Controls.Add(pnlTop)
        End Sub

        ''' <summary>DataGridView の DoubleBuffered プロパティをリフレクションで有効化する。</summary>
        Private Shared Sub EnableDoubleBuffering(control As Control)
            Dim prop As System.Reflection.PropertyInfo = GetType(Control).GetProperty(
                "DoubleBuffered",
                System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
            If prop IsNot Nothing Then
                prop.SetValue(control, True, Nothing)
            End If
        End Sub

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

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If _filterTimer IsNot Nothing Then _filterTimer.Dispose()
                If _bindingSource IsNot Nothing Then _bindingSource.Dispose()
                If _dataTable IsNot Nothing Then _dataTable.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

    End Class

End Namespace
