Option Explicit On
Option Strict On
Option Infer Off

Imports System.IO
Imports System.Text

Namespace Services

    ''' <summary>
    ''' アプリケーション全体で使用するシンプルなファイルロガー。
    ''' exe と同じフォルダに OutlookVault_yyyyMMdd.log を出力する。
    ''' スレッドセーフ。
    ''' </summary>
    Public Class Logger

        Private Shared ReadOnly _lock As New Object()
        Private Shared _logDirectory As String

        ''' <summary>ログ出力先ディレクトリ。テスト用にオーバーライド可能。</summary>
        Public Shared Property LogDirectory As String
            Get
                If _logDirectory Is Nothing Then
                    _logDirectory = Path.GetDirectoryName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location)
                End If
                Return _logDirectory
            End Get
            Set(value As String)
                _logDirectory = value
            End Set
        End Property

        ''' <summary>情報レベルのログを出力する。</summary>
        Public Shared Sub Info(message As String)
            WriteLog("INFO", message)
        End Sub

        ''' <summary>警告レベルのログを出力する。</summary>
        Public Shared Sub Warn(message As String)
            WriteLog("WARN", message)
        End Sub

        ''' <summary>エラーレベルのログを出力する。</summary>
        Public Shared Sub [Error](message As String)
            WriteLog("ERROR", message)
        End Sub

        ''' <summary>例外付きのエラーレベルのログを出力する。スタックトレースも記録する。</summary>
        Public Shared Sub [Error](message As String, ex As Exception)
            WriteLog("ERROR", message & " - " & ex.ToString())
        End Sub

        ''' <summary>現在の日付に対応するログファイルのフルパスを返す。</summary>
        Public Shared Function GetLogFilePath() As String
            Dim fileName As String = "OutlookVault_" & DateTime.Now.ToString("yyyyMMdd") & ".log"
            Return Path.Combine(LogDirectory, fileName)
        End Function

        Private Shared Sub WriteLog(level As String, message As String)
            SyncLock _lock
                Try
                    Dim logPath As String = GetLogFilePath()
                    Dim line As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                        " [" & level & "] " & message
                    Using sw As New StreamWriter(logPath, True, Encoding.UTF8)
                        sw.WriteLine(line)
                    End Using
                Catch ex As Exception
                    ' ログ出力自体の失敗はアプリを止めない。デバッグ時の手がかりとして出力ウィンドウに記録する。
                    System.Diagnostics.Debug.WriteLine("Logger write failed: " & ex.Message)
                End Try
            End SyncLock
        End Sub

    End Class

End Namespace
