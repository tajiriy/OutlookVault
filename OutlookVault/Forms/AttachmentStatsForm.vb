Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.SQLite
Imports System.Windows.Forms

Namespace Forms

    ''' <summary>
    ''' 添付ファイルの拡張子別統計情報を表示するフォーム。
    ''' 拡張子ごとの件数・合計サイズ・割合を一覧表示する。
    ''' </summary>
    Partial Public Class AttachmentStatsForm

        Private ReadOnly _dbManager As Data.DatabaseManager
        Private _dataTable As System.Data.DataTable
        Private _bindingSource As BindingSource

        Public Sub New(dbManager As Data.DatabaseManager)
            _dbManager = dbManager
            _bindingSource = New BindingSource()
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon
            AddHandler dgv.CellFormatting, AddressOf Dgv_CellFormatting
            LoadData()
        End Sub

        Private Sub LoadData()
            Dim totalCount As Long = 0
            Dim totalSize As Long = 0
            Dim extMap As New System.Collections.Generic.Dictionary(Of String, ExtensionRow)()

            ' SQL 側で拡張子ごとに集計（全行転送を回避）
            Const sql As String =
                "SELECT " &
                "  CASE WHEN INSTR(file_name, '.') > 0 " &
                "    THEN LOWER(SUBSTR(file_name, LENGTH(file_name) - LENGTH(SUBSTR(file_name, INSTR(file_name, '.'))) + 1)) " &
                "    ELSE '' END AS ext, " &
                "  COUNT(*) AS cnt, " &
                "  SUM(COALESCE(file_size, 0)) AS total_size " &
                "FROM attachments GROUP BY ext ORDER BY cnt DESC"
            Using conn As SQLiteConnection = _dbManager.GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim ext As String = reader.GetString(0)
                            If String.IsNullOrEmpty(ext) Then ext = "(なし)"
                            Dim cnt As Long = reader.GetInt64(1)
                            Dim size As Long = reader.GetInt64(2)
                            extMap(ext) = New ExtensionRow(ext, cnt, size)
                            totalCount += cnt
                            totalSize += size
                        End While
                    End Using
                End Using
            End Using

            Dim rows As New System.Collections.Generic.List(Of ExtensionRow)(extMap.Values)
            rows.Sort(Function(a, b) b.Count.CompareTo(a.Count))

            dgv.Rows.Clear()
            For Each row As ExtensionRow In rows
                Dim pct As Double = If(totalCount > 0, row.Count * 100.0 / CDbl(totalCount), 0.0)
                Dim icon As System.Drawing.Bitmap = Icons.ShellIconHelper.GetExtensionIconSmall(row.Extension)
                dgv.Rows.Add(icon, row.Extension, row.Count, row.TotalSize, pct)
            Next

            lblSummary.Text = String.Format("合計: {0:#,##0} 件 / {1}  ({2} 種類)",
                                            totalCount, FormatFileSize(totalSize), rows.Count)
        End Sub

        Private Sub Dgv_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
            If e.ColumnIndex = colTotalSize.Index AndAlso e.Value IsNot Nothing Then
                Dim bytes As Long = CLng(e.Value)
                e.Value = FormatFileSize(bytes)
                e.FormattingApplied = True
            End If
        End Sub

        Private Shared Function FormatFileSize(bytes As Long) As String
            Return Services.FileHelper.FormatFileSize(bytes)
        End Function

        Private Structure ExtensionRow
            Public ReadOnly Extension As String
            Public ReadOnly Count As Long
            Public ReadOnly TotalSize As Long

            Public Sub New(ext As String, cnt As Long, size As Long)
                Extension = ext
                Count = cnt
                TotalSize = size
            End Sub
        End Structure

        Private Sub AttachmentStatsForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
            If _bindingSource IsNot Nothing Then _bindingSource.Dispose()
            If _dataTable IsNot Nothing Then _dataTable.Dispose()
        End Sub

    End Class

End Namespace
