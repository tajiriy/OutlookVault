Option Explicit On
Option Strict On
Option Infer Off

Imports System.IO
Imports System.Windows.Forms

Namespace Forms

    Partial Public Class HelpForm

        Public Sub New()
            InitializeComponent()
            Dim appIcon As Drawing.Icon = Services.FileHelper.GetAppIcon()
            If appIcon IsNot Nothing Then Me.Icon = appIcon

            Dim helpPath As String = Path.Combine(Application.StartupPath, "Help", "user-manual.html")
            If File.Exists(helpPath) Then
                webBrowser.Navigate(helpPath)
            Else
                webBrowser.DocumentText = "<html><body style='font-family:Meiryo UI;padding:32px;'>" &
                    "<h2>ユーザーマニュアルが見つかりません</h2>" &
                    "<p>以下のパスにファイルが存在しません:</p>" &
                    "<pre>" & helpPath & "</pre></body></html>"
            End If
        End Sub

    End Class

End Namespace
