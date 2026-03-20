Option Explicit On
Option Strict On
Option Infer Off

Imports System.Windows.Forms

Namespace Forms

    ''' <summary>自動削除ルールの追加・編集ダイアログ。</summary>
    Partial Public Class AutoDeleteRuleEditDialog

        Private ReadOnly _emailRepo As Data.EmailRepository

        ''' <summary>入力されたルール名。</summary>
        Public Property RuleName As String

        ''' <summary>入力されたフィルタ式。</summary>
        Public Property FilterExpression As String

        Public Sub New(emailRepo As Data.EmailRepository,
                       Optional ruleName As String = "",
                       Optional filterExpression As String = "")
            _emailRepo = emailRepo
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon
            txtName.Text = ruleName
            txtFilter.Text = filterExpression
        End Sub

        Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
            If String.IsNullOrWhiteSpace(txtFilter.Text) Then
                lblPreviewResult.Text = "フィルタ式を入力してください"
                Return
            End If

            Try
                Dim matched As List(Of Models.Email) = _emailRepo.SearchEmailsFiltered(txtFilter.Text.Trim())
                lblPreviewResult.Text = String.Format("マッチ: {0} 件", matched.Count)
            Catch ex As Exception
                lblPreviewResult.Text = "エラー: " & ex.Message
            End Try
        End Sub

        Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
            If String.IsNullOrWhiteSpace(txtName.Text) Then
                MessageBox.Show("ルール名を入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtName.Focus()
                Return
            End If

            If String.IsNullOrWhiteSpace(txtFilter.Text) Then
                MessageBox.Show("フィルタ式を入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFilter.Focus()
                Return
            End If

            RuleName = txtName.Text.Trim()
            FilterExpression = txtFilter.Text.Trim()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

    End Class

End Namespace
