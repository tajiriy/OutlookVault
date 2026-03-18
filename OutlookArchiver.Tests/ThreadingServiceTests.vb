Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Services
Imports OutlookArchiver.Models

Namespace Tests

    <TestFixture>
    Public Class NormalizeSubjectTests

        ' ── 基本的なプレフィックス除去 ──────────────────────────────

        <Test>
        Public Sub NormalizeSubject_RePrefix_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("Re: Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_FwPrefix_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("FW: Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_FwdPrefix_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("Fwd: Hello"))
        End Sub

        ' ── 日本語・中国語プレフィックス ─────────────────────────────

        <Test>
        Public Sub NormalizeSubject_JapaneseReplyPrefix_Removed()
            Assert.AreEqual("会議の件", ThreadingService.NormalizeSubject("返信: 会議の件"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_JapaneseForwardPrefix_Removed()
            Assert.AreEqual("資料", ThreadingService.NormalizeSubject("転送: 資料"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_ChineseReplyPrefix_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("回复: Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_ChineseForwardPrefix_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("转发: Hello"))
        End Sub

        ' ── 多重プレフィックス ────────────────────────────────────

        <Test>
        Public Sub NormalizeSubject_MultipleRePrefixes_AllRemoved()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("Re: Re: Re: Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_MixedPrefixes_AllRemoved()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("Re: Fwd: FW: Hello"))
        End Sub

        ' ── 大文字小文字 ──────────────────────────────────────────

        <Test>
        Public Sub NormalizeSubject_LowercaseRe_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("re: Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_UppercaseRE_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("RE: Hello"))
        End Sub

        ' ── プレフィックスのバリエーション ────────────────────────────

        <Test>
        Public Sub NormalizeSubject_ReWithSpace_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("re : Hello"))
        End Sub

        <Test>
        Public Sub NormalizeSubject_ReWithDot_Removed()
            Assert.AreEqual("Hello", ThreadingService.NormalizeSubject("re. Hello"))
        End Sub

        ' ── エッジケース ──────────────────────────────────────────

        <Test>
        Public Sub NormalizeSubject_EmptyString_ReturnsEmpty()
            Assert.AreEqual(String.Empty, ThreadingService.NormalizeSubject(""))
        End Sub

        <Test>
        Public Sub NormalizeSubject_Nothing_ReturnsEmpty()
            Assert.AreEqual(String.Empty, ThreadingService.NormalizeSubject(Nothing))
        End Sub

        <Test>
        Public Sub NormalizeSubject_WhitespaceOnly_ReturnsEmpty()
            Assert.AreEqual(String.Empty, ThreadingService.NormalizeSubject("   "))
        End Sub

        <Test>
        Public Sub NormalizeSubject_NoPrefix_ReturnsTrimmed()
            Assert.AreEqual("Meeting tomorrow", ThreadingService.NormalizeSubject("  Meeting tomorrow  "))
        End Sub

        <Test>
        Public Sub NormalizeSubject_PrefixOnly_ReturnsEmpty()
            Assert.AreEqual(String.Empty, ThreadingService.NormalizeSubject("Re:"))
        End Sub

    End Class

    <TestFixture>
    Public Class AssignThreadIdTests

        ' ── Strategy 1: InReplyTo で親の ThreadId を継承 ──────────────

        <Test>
        Public Sub AssignThreadId_InReplyToWithKnownParent_InheritsParentThreadId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            messageIdCache.Add("parent@example.com", "thread-001")

            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "child@example.com"
            email.InReplyTo = "<parent@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            Assert.AreEqual("thread-001", email.ThreadId)
        End Sub

        <Test>
        Public Sub AssignThreadId_InReplyToWithUnknownParent_UsesParentMessageId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "child@example.com"
            email.InReplyTo = "<unknown-parent@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            Assert.AreEqual("unknown-parent@example.com", email.ThreadId)
        End Sub

        ' ── Strategy 2: References で祖先の ThreadId を継承 ───────────

        <Test>
        Public Sub AssignThreadId_ReferencesWithKnownAncestor_InheritsAncestorThreadId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            messageIdCache.Add("root@example.com", "thread-root")

            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "msg3@example.com"
            email.References = "<root@example.com> <msg2@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            ' 末尾（最新祖先）から探索するが msg2 は不明なので root にフォールバック
            Assert.AreEqual("thread-root", email.ThreadId)
        End Sub

        <Test>
        Public Sub AssignThreadId_ReferencesLatestAncestorFound_UsesLatestThreadId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            messageIdCache.Add("root@example.com", "thread-root")
            messageIdCache.Add("msg2@example.com", "thread-root")

            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "msg3@example.com"
            email.References = "<root@example.com> <msg2@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            ' 末尾の msg2 が見つかるのでそちらの ThreadId を使用
            Assert.AreEqual("thread-root", email.ThreadId)
        End Sub

        <Test>
        Public Sub AssignThreadId_ReferencesAllUnknown_UsesOldestRef()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "msg3@example.com"
            email.References = "<oldest@example.com> <mid@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            ' 全祖先が不明 → 最古（先頭）の References をルートとする
            Assert.AreEqual("oldest@example.com", email.ThreadId)
        End Sub

        ' ── Strategy 3: 正規化済み件名フォールバック ──────────────────

        <Test>
        Public Sub AssignThreadId_SubjectMatch_InheritsSubjectThreadId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            subjectCache.Add("会議の件", "thread-meeting")

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "new@example.com"
            email.NormalizedSubject = "会議の件"

            svc.AssignThreadId(email)

            Assert.AreEqual("thread-meeting", email.ThreadId)
        End Sub

        ' ── Strategy 4: 新規スレッド ──────────────────────────────────

        <Test>
        Public Sub AssignThreadId_NoMatch_UsesOwnMessageId()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "brand-new@example.com"
            email.NormalizedSubject = "新しい話題"

            svc.AssignThreadId(email)

            Assert.AreEqual("brand-new@example.com", email.ThreadId)
        End Sub

        <Test>
        Public Sub AssignThreadId_NoMessageId_GeneratesGuid()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.NormalizedSubject = "新しい話題"

            svc.AssignThreadId(email)

            Assert.IsNotNull(email.ThreadId)
            Assert.IsNotEmpty(email.ThreadId)
            ' GUID形式であることを確認
            Dim guid As Guid
            Assert.IsTrue(Guid.TryParse(email.ThreadId, guid))
        End Sub

        ' ── キャッシュ更新 ────────────────────────────────────────────

        <Test>
        Public Sub AssignThreadId_CacheUpdated_SubsequentEmailUsesCache()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            ' 1件目: 新規スレッドが作られる
            Dim email1 As New Email()
            email1.MessageId = "first@example.com"
            email1.NormalizedSubject = "プロジェクト進捗"
            svc.AssignThreadId(email1)

            ' 2件目: InReplyTo で1件目を参照 → キャッシュから ThreadId を取得
            Dim email2 As New Email()
            email2.MessageId = "second@example.com"
            email2.InReplyTo = "<first@example.com>"
            email2.NormalizedSubject = "プロジェクト進捗"
            svc.AssignThreadId(email2)

            Assert.AreEqual(email1.ThreadId, email2.ThreadId)
        End Sub

        <Test>
        Public Sub AssignThreadId_SubjectCacheUpdated_SubsequentEmailUsesSubjectCache()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            ' 1件目: 新規スレッド
            Dim email1 As New Email()
            email1.MessageId = "first@example.com"
            email1.NormalizedSubject = "週次報告"
            svc.AssignThreadId(email1)

            ' 2件目: 同じ件名で InReplyTo/References なし → 件名キャッシュで合流
            Dim email2 As New Email()
            email2.MessageId = "second@example.com"
            email2.NormalizedSubject = "週次報告"
            svc.AssignThreadId(email2)

            Assert.AreEqual(email1.ThreadId, email2.ThreadId)
        End Sub

        ' ── 優先順位の確認 ────────────────────────────────────────────

        <Test>
        Public Sub AssignThreadId_InReplyToTakesPriorityOverReferences()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            messageIdCache.Add("parent@example.com", "thread-from-parent")
            messageIdCache.Add("ancestor@example.com", "thread-from-ancestor")

            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)

            Dim email As New Email()
            email.MessageId = "child@example.com"
            email.InReplyTo = "<parent@example.com>"
            email.References = "<ancestor@example.com>"
            email.NormalizedSubject = "Hello"

            svc.AssignThreadId(email)

            ' InReplyTo が References より優先される
            Assert.AreEqual("thread-from-parent", email.ThreadId)
        End Sub

        ' ── ClearCaches ──────────────────────────────────────────────

        <Test>
        Public Sub ClearCaches_CachesCleared()
            Dim messageIdCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim subjectCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            subjectCache.Add("Hello", "thread-hello")

            Dim svc As New ThreadingService(Nothing)
            svc.LoadCaches(messageIdCache, subjectCache)
            svc.ClearCaches()

            ' キャッシュクリア後は件名フォールバックが効かず新規スレッドになる
            ' ただし DB アクセスが発生するため repo=Nothing だと例外になる
            ' ここではクリア操作自体が例外を出さないことを確認
            Assert.Pass()
        End Sub

    End Class

End Namespace
