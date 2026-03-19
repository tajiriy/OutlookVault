Option Explicit On
Option Strict On
Option Infer Off

Namespace Services

    ''' <summary>
    ''' メールのスレッド ID 付与と件名正規化を担当するサービスクラス。
    ''' </summary>
    Public Class ThreadingService

        Private ReadOnly _repo As Data.EmailRepository

        ' ── インメモリキャッシュ（取り込みセッション中のみ有効）────────
        Private _messageIdCache As Dictionary(Of String, String) = Nothing
        Private _subjectCache As Dictionary(Of String, String) = Nothing

        Public Sub New(repo As Data.EmailRepository)
            _repo = repo
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  キャッシュ管理（取り込みセッション用）
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 取り込み開始前に呼び出す。DB 全件を Dictionary にロードし、
        ''' セッション中の AssignThreadId が DB クエリなしで動作できるようにする。
        ''' </summary>
        Public Sub LoadCaches(messageIdMap As Dictionary(Of String, String),
                              subjectMap As Dictionary(Of String, String))
            _messageIdCache = messageIdMap
            _subjectCache = subjectMap
        End Sub

        ''' <summary>取り込みセッション終了後にキャッシュを解放する。</summary>
        Public Sub ClearCaches()
            _messageIdCache = Nothing
            _subjectCache = Nothing
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  スレッド ID 付与
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' email.ThreadId を決定する。以下の優先順位で判定する。
        ''' 1. In-Reply-To ヘッダー → 親メールの ThreadId を継承
        ''' 2. References ヘッダー → 祖先メールの ThreadId を継承
        ''' 3. 正規化済み件名フォールバック → 同名スレッドに合流
        ''' 4. 新規スレッド → email.MessageId を ThreadId として使用
        ''' キャッシュがロードされている場合は DB クエリの代わりにインメモリ参照を使用する。
        ''' </summary>
        Public Sub AssignThreadId(email As Models.Email)
            AssignThreadIdCore(email)
            ' 割り当て結果をキャッシュに追記（同セッション内の後続メールが参照できるよう）
            UpdateCaches(email)
        End Sub

        Private Sub AssignThreadIdCore(email As Models.Email)

            ' ── Strategy 1: In-Reply-To ─────────────────────────────
            If Not String.IsNullOrEmpty(email.InReplyTo) Then
                Dim parentId As String = CleanMessageId(email.InReplyTo)
                Dim parentThreadId As String = ResolveThreadIdByMessageId(parentId)
                If Not String.IsNullOrEmpty(parentThreadId) Then
                    email.ThreadId = parentThreadId
                    Return
                End If
                ' 親がまだ取り込まれていない場合は親の MessageId をスレッドルートとする
                email.ThreadId = parentId
                Return
            End If

            ' ── Strategy 2: References ──────────────────────────────
            If Not String.IsNullOrEmpty(email.References) Then
                Dim refs As String() = ParseReferences(email.References)
                ' 最新の祖先から順に検索する（References は古い順に並ぶため末尾から）
                Dim i As Integer = refs.Length - 1
                Do While i >= 0
                    Dim refId As String = CleanMessageId(refs(i))
                    Dim ancestorThreadId As String = ResolveThreadIdByMessageId(refId)
                    If Not String.IsNullOrEmpty(ancestorThreadId) Then
                        email.ThreadId = ancestorThreadId
                        Return
                    End If
                    i -= 1
                Loop
                ' DB に祖先がなければ最古の References をルートとする
                If refs.Length > 0 Then
                    email.ThreadId = CleanMessageId(refs(0))
                    Return
                End If
            End If

            ' ── Strategy 3: 正規化済み件名フォールバック ─────────────
            If Not String.IsNullOrEmpty(email.NormalizedSubject) Then
                Dim subjectThreadId As String = ResolveThreadIdBySubject(email.NormalizedSubject)
                If Not String.IsNullOrEmpty(subjectThreadId) Then
                    email.ThreadId = subjectThreadId
                    Return
                End If
            End If

            ' ── Strategy 4: 新規スレッド ────────────────────────────
            email.ThreadId = If(Not String.IsNullOrEmpty(email.MessageId),
                                email.MessageId,
                                System.Guid.NewGuid().ToString())
        End Sub

        ''' <summary>MessageId から ThreadId を解決する。キャッシュ優先、なければ DB。</summary>
        Private Function ResolveThreadIdByMessageId(messageId As String) As String
            If _messageIdCache IsNot Nothing Then
                Dim cached As String = Nothing
                If _messageIdCache.TryGetValue(messageId, cached) Then Return cached
                Return Nothing  ' キャッシュにない = 存在しない
            End If
            ' キャッシュなし → DB クエリ
            Dim found As Models.Email = _repo.GetEmailByMessageId(messageId)
            If found IsNot Nothing Then Return found.ThreadId
            Return Nothing
        End Function

        ''' <summary>正規化済み件名から ThreadId を解決する。キャッシュ優先、なければ DB。</summary>
        Private Function ResolveThreadIdBySubject(normalizedSubject As String) As String
            If _subjectCache IsNot Nothing Then
                Dim cached As String = Nothing
                If _subjectCache.TryGetValue(normalizedSubject, cached) Then Return cached
                Return Nothing
            End If
            ' キャッシュなし → DB クエリ
            Dim related As List(Of Models.Email) = _repo.GetEmailsByNormalizedSubject(normalizedSubject, 1)
            If related.Count > 0 Then Return related(0).ThreadId
            Return Nothing
        End Function

        ''' <summary>割り当て済みスレッド ID をキャッシュに追記する。</summary>
        Private Sub UpdateCaches(email As Models.Email)
            If String.IsNullOrEmpty(email.ThreadId) Then Return
            If _messageIdCache IsNot Nothing AndAlso Not String.IsNullOrEmpty(email.MessageId) Then
                If Not _messageIdCache.ContainsKey(email.MessageId) Then
                    _messageIdCache.Add(email.MessageId, email.ThreadId)
                End If
            End If
            If _subjectCache IsNot Nothing AndAlso Not String.IsNullOrEmpty(email.NormalizedSubject) Then
                If Not _subjectCache.ContainsKey(email.NormalizedSubject) Then
                    _subjectCache.Add(email.NormalizedSubject, email.ThreadId)
                End If
            End If
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  件名正規化
        ' ════════════════════════════════════════════════════════════

        ''' <summary>
        ''' 返信・転送プレフィックスを除去した正規化済み件名を返す。
        ''' 除去対象: Re: / FW: / Fwd: / 返信: / 転送: およびその変形。
        ''' </summary>
        Public Shared Function NormalizeSubject(subject As String) As String
            If String.IsNullOrWhiteSpace(subject) Then Return String.Empty

            Dim result As String = subject.Trim()
            Dim prefixes As String() = {
                "re:", "re :", "re. ",
                "fw:", "fw :", "fw. ",
                "fwd:", "fwd :", "fwd. ",
                "返信:", "転送:", "回复:", "转发:"
            }

            Dim changed As Boolean = True
            Do While changed
                changed = False
                For Each prefix As String In prefixes
                    If result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                        result = result.Substring(prefix.Length).TrimStart()
                        changed = True
                    End If
                Next
            Loop

            Return result
        End Function

        ' ════════════════════════════════════════════════════════════
        '  ヘルパー
        ' ════════════════════════════════════════════════════════════

        ''' <summary>MessageID の前後の山括弧を除去する。</summary>
        Private Shared Function CleanMessageId(messageId As String) As String
            If String.IsNullOrEmpty(messageId) Then Return messageId
            Dim result As String = messageId.Trim()
            If result.StartsWith("<") AndAlso result.EndsWith(">") Then
                result = result.Substring(1, result.Length - 2)
            End If
            Return result
        End Function

        ''' <summary>References ヘッダー値をスペース区切りで分割して配列にする。</summary>
        Private Shared Function ParseReferences(references As String) As String()
            Return references.Split(New Char() {" "c, Chr(9), Chr(13), Chr(10)},
                                    StringSplitOptions.RemoveEmptyEntries)
        End Function

    End Class

End Namespace
