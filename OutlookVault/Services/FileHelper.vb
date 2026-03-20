Option Explicit On
Option Strict On
Option Infer Off

Namespace Services

    ''' <summary>
    ''' ファイル関連のユーティリティメソッドを提供する。
    ''' </summary>
    Public Class FileHelper

        ''' <summary>バイト数を人間可読な形式（B / KB / MB / GB）に変換する。</summary>
        Public Shared Function FormatFileSize(bytes As Long) As String
            If bytes < 1024L Then
                Return bytes.ToString("#,##0") & " B"
            ElseIf bytes < 1024L * 1024L Then
                Return (bytes / 1024.0).ToString("#,##0.0") & " KB"
            ElseIf bytes < 1024L * 1024L * 1024L Then
                Return (bytes / (1024.0 * 1024.0)).ToString("#,##0.0") & " MB"
            Else
                Return (bytes / (1024.0 * 1024.0 * 1024.0)).ToString("#,##0.0") & " GB"
            End If
        End Function

        Private Shared _appIcon As Drawing.Icon

        ''' <summary>アプリケーションアイコン (app.ico) を返す。見つからなければ Nothing。</summary>
        Public Shared Function GetAppIcon() As Drawing.Icon
            If _appIcon Is Nothing Then
                Dim icoPath As String = System.IO.Path.Combine(Application.StartupPath, "app.ico")
                If System.IO.File.Exists(icoPath) Then
                    _appIcon = New Drawing.Icon(icoPath)
                End If
            End If
            Return _appIcon
        End Function

    End Class

End Namespace
