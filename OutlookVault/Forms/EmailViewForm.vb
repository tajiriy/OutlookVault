Option Explicit On
Option Strict On
Option Infer Off

Namespace Forms

    ''' <summary>
    ''' メールを別ウィンドウで表示するフォーム。
    ''' メール一覧のダブルクリックで開かれる。
    ''' </summary>
    Public Class EmailViewForm

        Public Sub New()
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon
        End Sub

    ''' <summary>指定メールをプレビューに表示する。</summary>
    Public Sub ShowEmail(email As Models.Email, Optional highlightQuery As String = Nothing)
        ' タイトルバーに件名を表示
        Dim subject As String = If(String.IsNullOrEmpty(email.Subject), "(件名なし)", email.Subject)
        Me.Text = subject & " - OutlookVault"

        emailPreview.ShowEmail(email, highlightQuery)

        ' HTML/テキスト切り替えボタンの表示制御
        UpdateToggleButton()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  イベントハンドラ
    ' ════════════════════════════════════════════════════════════

    Private Sub btnToggleView_Click(sender As Object, e As System.EventArgs) Handles btnToggleView.Click
        emailPreview.ToggleView()
        UpdateToggleButton()
    End Sub

    ' ════════════════════════════════════════════════════════════
    '  プライベートメソッド
    ' ════════════════════════════════════════════════════════════

    ''' <summary>HTML/テキスト切り替えボタンの表示状態とテキストを更新する。</summary>
    Private Sub UpdateToggleButton()
        btnToggleView.Visible = emailPreview.CanToggleView
        If emailPreview.IsHtmlView Then
            btnToggleView.Text = "テキスト表示"
        Else
            btnToggleView.Text = "HTML 表示"
        End If
    End Sub

    End Class

End Namespace
