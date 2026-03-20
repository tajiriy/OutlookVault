Option Explicit On
Option Strict On
Option Infer Off

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace Forms

    Partial Public Class AboutForm

        Public Sub New()
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon

            ' アイコン画像を読み込み
            Dim icoPath As String = System.IO.Path.Combine(Application.StartupPath, "app.ico")
            If System.IO.File.Exists(icoPath) Then
                Dim ico As New Icon(icoPath, 256, 256)
                picIcon.Image = ico.ToBitmap()
            End If

            ' バージョン情報を設定
            Dim ver As Version = Assembly.GetExecutingAssembly().GetName().Version
            lblVersion.Text = "Version " & ver.Major.ToString() & "." & ver.Minor.ToString() & "." & ver.Build.ToString()
        End Sub

        Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
            Me.Close()
        End Sub

    End Class

End Namespace
