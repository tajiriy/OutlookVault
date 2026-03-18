Option Explicit On
Option Strict On
Option Infer Off

Imports System.IO
Imports System.Text

Namespace Services

    ''' <summary>
    ''' 取り込みエラーの詳細をログファイルに出力するクラス。
    ''' exe と同じフォルダに import_errors_yyyyMMdd_HHmmss.log を生成する。
    ''' </summary>
    Public Class ImportLogWriter

        ''' <summary>
        ''' エラー詳細をログファイルに書き出し、ログファイルのフルパスを返す。
        ''' エラーが無い場合は Nothing を返す。
        ''' </summary>
        Public Shared Function WriteErrorLog(errors As List(Of ImportErrorEntry)) As String
            If errors Is Nothing OrElse errors.Count = 0 Then Return Nothing

            Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim fileName As String = "import_errors_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log"
            Dim filePath As String = Path.Combine(exeDir, fileName)

            Using sw As New StreamWriter(filePath, False, Encoding.UTF8)
                sw.WriteLine("OutlookArchiver 取り込みエラーログ")
                sw.WriteLine("出力日時: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                sw.WriteLine("エラー件数: " & errors.Count.ToString())
                sw.WriteLine(New String("="c, 80))
                sw.WriteLine()

                For idx As Integer = 0 To errors.Count - 1
                    Dim entry As ImportErrorEntry = errors(idx)
                    sw.WriteLine("[" & (idx + 1).ToString() & "] " & entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"))
                    sw.WriteLine("  フォルダ   : " & If(String.IsNullOrEmpty(entry.FolderName), "(不明)", entry.FolderName))
                    sw.WriteLine("  MessageID  : " & If(String.IsNullOrEmpty(entry.MessageId), "(不明)", entry.MessageId))
                    sw.WriteLine("  件名       : " & If(String.IsNullOrEmpty(entry.Subject), "(不明)", entry.Subject))
                    sw.WriteLine("  エラー内容 : " & entry.ErrorMessage)
                    sw.WriteLine()
                Next
            End Using

            Return filePath
        End Function

        ''' <summary>
        ''' エラーをエラーメッセージ別に集計し、メッセージと件数のペアを返す。
        ''' 件数の多い順にソートされる。
        ''' </summary>
        Public Shared Function SummarizeErrors(errors As List(Of ImportErrorEntry)) As List(Of KeyValuePair(Of String, Integer))
            Dim counts As New Dictionary(Of String, Integer)()
            For Each entry As ImportErrorEntry In errors
                Dim key As String = entry.ErrorMessage
                If counts.ContainsKey(key) Then
                    counts(key) += 1
                Else
                    counts(key) = 1
                End If
            Next

            Dim sorted As New List(Of KeyValuePair(Of String, Integer))(counts)
            sorted.Sort(Function(a As KeyValuePair(Of String, Integer), b As KeyValuePair(Of String, Integer)) b.Value.CompareTo(a.Value))
            Return sorted
        End Function

    End Class

End Namespace
