Option Explicit On
Option Strict On
Option Infer Off

Imports System.Text.RegularExpressions

Namespace Services

    ''' <summary>
    ''' メール本文から引用部分を除去するサービス。
    ''' 除去は表示時のみ行い、DB のオリジナル本文は変更しない。
    ''' </summary>
    Public Class QuoteStripperService

        ' HTML 引用除去用の事前コンパイル済み正規表現
        Private Shared ReadOnly BlockquoteRegex As New Regex(
            "<blockquote[^>]*>.*?</blockquote\s*>",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Private Shared ReadOnly GmailQuoteRegex As New Regex(
            "<div\s+class=""gmail_quote""[^>]*>.*?</div\s*>",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Private Shared ReadOnly OutlookReplyFwdRegex As New Regex(
            "<div\s+id=""divRplyFwdMsg"".*",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Private Shared ReadOnly OutlookMessageHeaderRegex As New Regex(
            "<div\s+[^>]*class=""[^""]*OutlookMessageHeader[^""]*"".*",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        ' 区切り線パターン（行全体または先頭一致）
        Private Shared ReadOnly SeparatorPatterns As String() = {
            "-----Original Message-----",
            "-----元のメッセージ-----",
            "-----Forwarded Message-----",
            "-----転送メッセージ-----",
            "________________________________",
            "------------------------------------------------------------------------"
        }

        ' ════════════════════════════════════════════════════════════
        '  公開メソッド
        ' ════════════════════════════════════════════════════════════

        ''' <summary>プレーンテキスト本文から引用部を除去する。</summary>
        Public Function StripQuotesFromText(text As String) As String
            If String.IsNullOrEmpty(text) Then Return text

            Dim lines() As String = text.Split(
                New String() {vbCrLf, vbLf, vbCr},
                StringSplitOptions.None)

            ' Step 1: セパレータ行以降を削除
            Dim cutIndex As Integer = FindSeparatorLine(lines)

            Dim workLines() As String
            If cutIndex >= 0 Then
                workLines = New String(cutIndex - 1) {}
                System.Array.Copy(lines, workLines, cutIndex)
            Else
                workLines = lines
            End If

            ' Step 2: 末尾から連続する空行・"> " 行を削除
            Dim lastIndex As Integer = TrimTrailingQuoteLines(workLines)

            Dim finalLines(lastIndex) As String
            System.Array.Copy(workLines, finalLines, lastIndex + 1)

            Return String.Join(vbCrLf, finalLines).TrimEnd()
        End Function

        ''' <summary>HTML 本文から引用ブロックを除去する。</summary>
        Public Function StripQuotesFromHtml(html As String) As String
            If String.IsNullOrEmpty(html) Then Return html

            Dim result As String = html

            ' blockquote 要素を削除（Outlook / メール全般の引用形式）
            result = BlockquoteRegex.Replace(result, String.Empty)

            ' Gmail 引用 div を削除
            result = GmailQuoteRegex.Replace(result, String.Empty)

            ' Outlook の返信/転送区切り div を削除（それ以降すべて）
            result = OutlookReplyFwdRegex.Replace(result, String.Empty)

            ' Outlook の OutlookMessageHeader div を削除（それ以降すべて）
            result = OutlookMessageHeaderRegex.Replace(result, String.Empty)

            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  プライベートメソッド
        ' ════════════════════════════════════════════════════════════

        ''' <summary>セパレータ行のインデックスを返す。見つからない場合は -1。</summary>
        Private Shared Function FindSeparatorLine(lines() As String) As Integer
            For i As Integer = 0 To lines.Length - 1
                Dim trimmed As String = lines(i).Trim()
                For Each pattern As String In SeparatorPatterns
                    If trimmed = pattern OrElse trimmed.StartsWith(pattern) Then
                        Return i
                    End If
                Next
            Next
            Return -1
        End Function

        ''' <summary>末尾から連続する空行・引用行（"> "）を除いた最後の有効インデックスを返す。</summary>
        Private Shared Function TrimTrailingQuoteLines(lines() As String) As Integer
            Dim lastIndex As Integer = lines.Length - 1
            While lastIndex >= 0
                Dim trimmedLine As String = lines(lastIndex).TrimEnd()
                If trimmedLine.Length = 0 OrElse trimmedLine.TrimStart().StartsWith(">") Then
                    lastIndex -= 1
                Else
                    Exit While
                End If
            End While
            Return lastIndex
        End Function

    End Class

End Namespace
