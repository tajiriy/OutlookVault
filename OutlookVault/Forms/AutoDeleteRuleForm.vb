Option Explicit On
Option Strict On
Option Infer Off

Imports System.Windows.Forms

Namespace Forms

    ''' <summary>自動削除ルールの管理フォーム。</summary>
    Partial Public Class AutoDeleteRuleForm

        Private ReadOnly _ruleRepo As Data.AutoDeleteRuleRepository
        Private ReadOnly _emailRepo As Data.EmailRepository
        Private ReadOnly _autoDeleteSvc As Services.AutoDeleteService
        Private _rules As List(Of Models.AutoDeleteRule)

        Public Sub New(ruleRepo As Data.AutoDeleteRuleRepository,
                       emailRepo As Data.EmailRepository,
                       autoDeleteSvc As Services.AutoDeleteService)
            _ruleRepo = ruleRepo
            _emailRepo = emailRepo
            _autoDeleteSvc = autoDeleteSvc
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon
            LoadRules()
        End Sub

        Private Sub LoadRules()
            _rules = _ruleRepo.GetAllRules()
            dgvRules.Rows.Clear()
            For Each rule As Models.AutoDeleteRule In _rules
                Dim idx As Integer = dgvRules.Rows.Add(rule.Name, rule.FilterExpression, rule.Enabled)
                dgvRules.Rows(idx).Tag = rule
            Next
            lblCount.Text = String.Format("{0} 件のルール", _rules.Count)
        End Sub

        Private Function GetSelectedRule() As Models.AutoDeleteRule
            If dgvRules.SelectedRows.Count = 0 Then Return Nothing
            Return TryCast(dgvRules.SelectedRows(0).Tag, Models.AutoDeleteRule)
        End Function

        Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
            Using dlg As New AutoDeleteRuleEditDialog(_emailRepo)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    _ruleRepo.InsertRule(dlg.RuleName, dlg.FilterExpression)
                    LoadRules()
                End If
            End Using
        End Sub

        Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
            Dim rule As Models.AutoDeleteRule = GetSelectedRule()
            If rule Is Nothing Then Return

            Using dlg As New AutoDeleteRuleEditDialog(_emailRepo, rule.Name, rule.FilterExpression)
                If dlg.ShowDialog(Me) = DialogResult.OK Then
                    _ruleRepo.UpdateRule(rule.Id, dlg.RuleName, dlg.FilterExpression, rule.Enabled)
                    LoadRules()
                End If
            End Using
        End Sub

        Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
            Dim rule As Models.AutoDeleteRule = GetSelectedRule()
            If rule Is Nothing Then Return

            If MessageBox.Show(
                    String.Format("ルール「{0}」を削除しますか？", rule.Name),
                    "ルールの削除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                Return
            End If

            _ruleRepo.DeleteRule(rule.Id)
            LoadRules()
        End Sub

        Private Sub btnToggle_Click(sender As Object, e As EventArgs) Handles btnToggle.Click
            Dim rule As Models.AutoDeleteRule = GetSelectedRule()
            If rule Is Nothing Then Return

            _ruleRepo.SetEnabled(rule.Id, Not rule.Enabled)
            LoadRules()
        End Sub

        Private Sub btnReapply_Click(sender As Object, e As EventArgs) Handles btnReapply.Click
            Dim enabledCount As Integer = 0
            For Each rule As Models.AutoDeleteRule In _rules
                If rule.Enabled Then enabledCount += 1
            Next

            If enabledCount = 0 Then
                MessageBox.Show("有効なルールがありません。", "再適用",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If MessageBox.Show(
                    String.Format("有効な {0} 件のルールを既存メールに再適用しますか？" & vbCrLf &
                                  "マッチしたメールはゴミ箱に移動します。", enabledCount),
                    "ルールの再適用",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then
                Return
            End If

            Dim deletedCount As Integer = _autoDeleteSvc.ApplyRulesToExistingEmails()
            MessageBox.Show(
                String.Format("再適用が完了しました。{0} 件のメールをゴミ箱に移動しました。", deletedCount),
                "再適用完了",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub dgvRules_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRules.CellDoubleClick
            If e.RowIndex < 0 Then Return
            btnEdit_Click(sender, e)
        End Sub

    End Class

End Namespace
