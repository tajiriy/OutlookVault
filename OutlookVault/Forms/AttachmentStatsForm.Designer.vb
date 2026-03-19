Option Explicit On
Option Strict On
Option Infer Off

Namespace Forms

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class AttachmentStatsForm
        Inherits System.Windows.Forms.Form

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing Then
                    If _dataTable IsNot Nothing Then _dataTable.Dispose()
                    If _bindingSource IsNot Nothing Then _bindingSource.Dispose()
                    If components IsNot Nothing Then components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        Private components As System.ComponentModel.IContainer

        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Me.pnlBottom = New System.Windows.Forms.Panel()
            Me.lblSummary = New System.Windows.Forms.Label()
            Me.btnClose = New System.Windows.Forms.Button()
            Me.dgv = New System.Windows.Forms.DataGridView()
            Me.colExtension = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colTotalSize = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colPercent = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.pnlBottom.SuspendLayout()
            CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'pnlBottom
            '
            Me.pnlBottom.Controls.Add(Me.lblSummary)
            Me.pnlBottom.Controls.Add(Me.btnClose)
            Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.pnlBottom.Location = New System.Drawing.Point(0, 371)
            Me.pnlBottom.Name = "pnlBottom"
            Me.pnlBottom.Padding = New System.Windows.Forms.Padding(8, 6, 8, 6)
            Me.pnlBottom.Size = New System.Drawing.Size(534, 40)
            Me.pnlBottom.TabIndex = 1
            '
            'lblSummary
            '
            Me.lblSummary.AutoSize = True
            Me.lblSummary.Dock = System.Windows.Forms.DockStyle.Left
            Me.lblSummary.Location = New System.Drawing.Point(8, 6)
            Me.lblSummary.Name = "lblSummary"
            Me.lblSummary.Size = New System.Drawing.Size(0, 12)
            Me.lblSummary.TabIndex = 0
            Me.lblSummary.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'btnClose
            '
            Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnClose.Dock = System.Windows.Forms.DockStyle.Right
            Me.btnClose.Location = New System.Drawing.Point(446, 6)
            Me.btnClose.Name = "btnClose"
            Me.btnClose.Size = New System.Drawing.Size(80, 28)
            Me.btnClose.TabIndex = 1
            Me.btnClose.Text = "閉じる"
            '
            'dgv
            '
            Me.dgv.AllowUserToAddRows = False
            Me.dgv.AllowUserToDeleteRows = False
            Me.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
            Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colExtension, Me.colCount, Me.colTotalSize, Me.colPercent})
            Me.dgv.Dock = System.Windows.Forms.DockStyle.Fill
            Me.dgv.Location = New System.Drawing.Point(0, 0)
            Me.dgv.Name = "dgv"
            Me.dgv.ReadOnly = True
            Me.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.dgv.Size = New System.Drawing.Size(534, 371)
            Me.dgv.TabIndex = 0
            '
            'colExtension
            '
            Me.colExtension.FillWeight = 30.0!
            Me.colExtension.HeaderText = "拡張子"
            Me.colExtension.Name = "colExtension"
            Me.colExtension.ReadOnly = True
            '
            'colCount
            '
            DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
            DataGridViewCellStyle1.Format = "#,##0"
            Me.colCount.DefaultCellStyle = DataGridViewCellStyle1
            Me.colCount.FillWeight = 20.0!
            Me.colCount.HeaderText = "件数"
            Me.colCount.Name = "colCount"
            Me.colCount.ReadOnly = True
            '
            'colTotalSize
            '
            DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
            Me.colTotalSize.DefaultCellStyle = DataGridViewCellStyle2
            Me.colTotalSize.FillWeight = 25.0!
            Me.colTotalSize.HeaderText = "合計サイズ"
            Me.colTotalSize.Name = "colTotalSize"
            Me.colTotalSize.ReadOnly = True
            '
            'colPercent
            '
            DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
            DataGridViewCellStyle3.Format = "0.0"
            Me.colPercent.DefaultCellStyle = DataGridViewCellStyle3
            Me.colPercent.FillWeight = 25.0!
            Me.colPercent.HeaderText = "割合 (%)"
            Me.colPercent.Name = "colPercent"
            Me.colPercent.ReadOnly = True
            '
            'AttachmentStatsForm
            '
            Me.AcceptButton = Me.btnClose
            Me.CancelButton = Me.btnClose
            Me.ClientSize = New System.Drawing.Size(534, 411)
            Me.Controls.Add(Me.dgv)
            Me.Controls.Add(Me.pnlBottom)
            Me.MinimumSize = New System.Drawing.Size(400, 300)
            Me.Name = "AttachmentStatsForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "添付ファイル拡張子統計"
            Me.pnlBottom.ResumeLayout(False)
            Me.pnlBottom.PerformLayout()
            CType(Me.dgv, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

        Friend pnlBottom As System.Windows.Forms.Panel
        Friend lblSummary As System.Windows.Forms.Label
        Friend WithEvents btnClose As System.Windows.Forms.Button
        Friend WithEvents dgv As System.Windows.Forms.DataGridView
        Friend colExtension As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend colCount As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend colTotalSize As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend colPercent As System.Windows.Forms.DataGridViewTextBoxColumn

    End Class

End Namespace
