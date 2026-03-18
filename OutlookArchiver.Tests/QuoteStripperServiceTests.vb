Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Services

Namespace Tests

    <TestFixture>
    Public Class StripQuotesFromTextTests

        Private _service As QuoteStripperService

        <SetUp>
        Public Sub SetUp()
            _service = New QuoteStripperService()
        End Sub

        ' ── セパレータ行による除去 ────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromText_OriginalMessageSeparator_RemovesAfterSeparator()
            Dim text As String =
                "こんにちは。" & vbCrLf &
                "よろしくお願いします。" & vbCrLf &
                "-----Original Message-----" & vbCrLf &
                "From: sender@example.com" & vbCrLf &
                "以前のメッセージ..."

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("こんにちは。"))
            Assert.IsTrue(result.Contains("よろしくお願いします。"))
            Assert.IsFalse(result.Contains("-----Original Message-----"))
            Assert.IsFalse(result.Contains("以前のメッセージ"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_JapaneseSeparator_RemovesAfterSeparator()
            Dim text As String =
                "了解しました。" & vbCrLf &
                "-----元のメッセージ-----" & vbCrLf &
                "元の内容です。"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("了解しました。"))
            Assert.IsFalse(result.Contains("元の内容です。"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_ForwardedMessageSeparator_RemovesAfterSeparator()
            Dim text As String =
                "FYI" & vbCrLf &
                "-----Forwarded Message-----" & vbCrLf &
                "転送されたメッセージ"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("FYI"))
            Assert.IsFalse(result.Contains("転送されたメッセージ"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_UnderscoreSeparator_RemovesAfterSeparator()
            Dim text As String =
                "本文" & vbCrLf &
                "________________________________" & vbCrLf &
                "引用部分"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("引用部分"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_DashSeparator_RemovesAfterSeparator()
            Dim text As String =
                "本文" & vbCrLf &
                "------------------------------------------------------------------------" & vbCrLf &
                "引用部分"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("引用部分"))
        End Sub

        ' ── 末尾引用行の除去 ──────────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromText_TrailingQuoteLines_Removed()
            Dim text As String =
                "返信内容です。" & vbCrLf &
                "" & vbCrLf &
                "> 元のメッセージ1" & vbCrLf &
                "> 元のメッセージ2" & vbCrLf &
                "> 元のメッセージ3"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("返信内容です。"))
            Assert.IsFalse(result.Contains("> 元のメッセージ"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_TrailingEmptyAndQuoteLines_Removed()
            Dim text As String =
                "本文" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf &
                "> 引用1" & vbCrLf &
                "" & vbCrLf &
                "> 引用2"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("引用"))
        End Sub

        ' ── エッジケース ──────────────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromText_NoQuotes_ReturnsOriginal()
            Dim text As String = "普通のメッセージです。" & vbCrLf & "2行目。"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("普通のメッセージです。"))
            Assert.IsTrue(result.Contains("2行目。"))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_EmptyString_ReturnsEmpty()
            Assert.AreEqual("", _service.StripQuotesFromText(""))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_Nothing_ReturnsNothing()
            Assert.IsNull(_service.StripQuotesFromText(Nothing))
        End Sub

        <Test>
        Public Sub StripQuotesFromText_SeparatorAndTrailingQuotes_BothRemoved()
            Dim text As String =
                "本文" & vbCrLf &
                "-----Original Message-----" & vbCrLf &
                "> 引用" & vbCrLf &
                "> 引用2"

            Dim result As String = _service.StripQuotesFromText(text)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("引用"))
        End Sub

    End Class

    <TestFixture>
    Public Class StripQuotesFromHtmlTests

        Private _service As QuoteStripperService

        <SetUp>
        Public Sub SetUp()
            _service = New QuoteStripperService()
        End Sub

        ' ── blockquote 除去 ──────────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromHtml_Blockquote_Removed()
            Dim html As String =
                "<p>返信です</p>" &
                "<blockquote><p>元のメッセージ</p></blockquote>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.IsTrue(result.Contains("返信です"))
            Assert.IsFalse(result.Contains("元のメッセージ"))
        End Sub

        <Test>
        Public Sub StripQuotesFromHtml_BlockquoteWithAttributes_Removed()
            Dim html As String =
                "<p>本文</p>" &
                "<blockquote style=""margin:0"" type=""cite""><p>引用</p></blockquote>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("引用"))
        End Sub

        ' ── Gmail 引用 div 除去 ─────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromHtml_GmailQuoteDiv_Removed()
            Dim html As String =
                "<p>返信内容</p>" &
                "<div class=""gmail_quote""><p>引用されたメッセージ</p></div>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.IsTrue(result.Contains("返信内容"))
            Assert.IsFalse(result.Contains("引用されたメッセージ"))
        End Sub

        ' ── Outlook 返信/転送 div 除去 ──────────────────────────────

        <Test>
        Public Sub StripQuotesFromHtml_OutlookRplyFwdMsg_RemovedWithEverythingAfter()
            Dim html As String =
                "<p>返信</p>" &
                "<div id=""divRplyFwdMsg""><p>From: sender</p></div>" &
                "<p>元のメッセージ全体</p>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.IsTrue(result.Contains("返信"))
            Assert.IsFalse(result.Contains("From: sender"))
            Assert.IsFalse(result.Contains("元のメッセージ全体"))
        End Sub

        <Test>
        Public Sub StripQuotesFromHtml_OutlookMessageHeader_RemovedWithEverythingAfter()
            Dim html As String =
                "<p>本文</p>" &
                "<div class=""OutlookMessageHeader""><p>ヘッダー情報</p></div>" &
                "<p>引用された本文</p>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.IsTrue(result.Contains("本文"))
            Assert.IsFalse(result.Contains("ヘッダー情報"))
            Assert.IsFalse(result.Contains("引用された本文"))
        End Sub

        ' ── エッジケース ──────────────────────────────────────────────

        <Test>
        Public Sub StripQuotesFromHtml_NoQuotes_ReturnsOriginal()
            Dim html As String = "<p>普通のメッセージ</p>"

            Dim result As String = _service.StripQuotesFromHtml(html)

            Assert.AreEqual(html, result)
        End Sub

        <Test>
        Public Sub StripQuotesFromHtml_EmptyString_ReturnsEmpty()
            Assert.AreEqual("", _service.StripQuotesFromHtml(""))
        End Sub

        <Test>
        Public Sub StripQuotesFromHtml_Nothing_ReturnsNothing()
            Assert.IsNull(_service.StripQuotesFromHtml(Nothing))
        End Sub

    End Class

End Namespace
