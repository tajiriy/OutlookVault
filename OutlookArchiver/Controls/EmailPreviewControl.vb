Option Explicit On
Option Strict On
Option Infer Off

Imports System.IO
Imports System.Text.RegularExpressions

Namespace Controls

    ''' <summary>
    ''' メールプレビュー UserControl。
    ''' ヘッダー（差出人・日時・件名・宛先）、本文（HTML/テキスト切り替え）、
    ''' 添付ファイルパネルを提供する。
    ''' </summary>
    Public Class EmailPreviewControl
        Inherits System.Windows.Forms.UserControl

        Private _currentEmail As Models.Email
        Private _showHtml As Boolean
        Private _highlightQuery As String

        ' ════════════════════════════════════════════════════════════
        '  初期化
        ' ════════════════════════════════════════════════════════════

        Public Sub New()
            InitializeComponent()
            ' キャプションラベルにボールドフォントを適用
            Dim boldFont As New System.Drawing.Font(Me.Font, System.Drawing.FontStyle.Bold)
            lblFromCaption.Font = boldFont
            lblDateCaption.Font = boldFont
            lblSubjectCaption.Font = boldFont
            lblToCaption.Font = boldFont
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  公開メソッド
        ' ════════════════════════════════════════════════════════════

        ''' <summary>指定メールの内容をプレビューに表示する。添付ファイル一覧も含む。</summary>
        ''' <param name="highlightQuery">HTML 表示時にハイライトする検索クエリ。Nothing の場合はハイライトなし。</param>
        Public Sub ShowEmail(email As Models.Email, Optional highlightQuery As String = Nothing)
            _currentEmail = email
            _highlightQuery = highlightQuery
            UpdateHeader(email)
            UpdateBody(email)
            LoadAttachments(email.Attachments)
        End Sub

        ''' <summary>プレビュー内容をクリアし、空の状態に戻す。</summary>
        Public Sub ClearPreview()
            _currentEmail = Nothing
            lblFromValue.Text = String.Empty
            lblDateValue.Text = String.Empty
            lblSubjectValue.Text = String.Empty
            lblToValue.Text = String.Empty
            webBrowser.DocumentText = "<html><body></body></html>"
            txtBodyText.Text = String.Empty
            flowAttachments.Controls.Clear()
            pnlAttachments.Visible = False
            btnToggleView.Enabled = False
            btnToggleView.Text = "テキスト表示"
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  プライベートメソッド
        ' ════════════════════════════════════════════════════════════

        Private Sub UpdateHeader(email As Models.Email)
            Dim senderName As String = If(String.IsNullOrEmpty(email.SenderName), String.Empty, email.SenderName)
            Dim senderMail As String = If(String.IsNullOrEmpty(email.SenderEmail), String.Empty, email.SenderEmail)
            If senderName.Length > 0 AndAlso senderMail.Length > 0 Then
                lblFromValue.Text = senderName & " <" & senderMail & ">"
            ElseIf senderMail.Length > 0 Then
                lblFromValue.Text = senderMail
            Else
                lblFromValue.Text = senderName
            End If

            lblDateValue.Text = email.ReceivedAt.ToString("yyyy/MM/dd HH:mm")
            lblSubjectValue.Text = If(String.IsNullOrEmpty(email.Subject), "(件名なし)", email.Subject)
            lblToValue.Text = FormatRecipientsJson(email.ToRecipients)
        End Sub

        Private Sub UpdateBody(email As Models.Email)
            Dim hasHtml As Boolean = Not String.IsNullOrEmpty(email.BodyHtml)
            Dim hasText As Boolean = Not String.IsNullOrEmpty(email.BodyText)

            ' 両方ある場合のみトグル有効
            btnToggleView.Enabled = hasHtml AndAlso hasText
            _showHtml = hasHtml

            If _showHtml Then
                Dim html As String = If(email.BodyHtml, String.Empty)
                html = ReplaceCidReferences(html, email.Attachments)
                If Not String.IsNullOrEmpty(_highlightQuery) Then
                    html = InjectHighlightScript(html, _highlightQuery)
                End If
                webBrowser.DocumentText = html
                webBrowser.Visible = True
                txtBodyText.Visible = False
                btnToggleView.Text = "テキスト表示"
            Else
                txtBodyText.Text = If(email.BodyText, String.Empty)
                txtBodyText.Visible = True
                webBrowser.Visible = False
                btnToggleView.Text = "HTML 表示"
            End If
        End Sub

        ''' <summary>
        ''' HTML 中の cid:xxx 参照を、対応するインライン画像の file:/// URL に置換して返す。
        ''' 対応する添付がない cid: はそのまま残す。
        ''' </summary>
        Private Shared Function ReplaceCidReferences(html As String,
                                                     attachments As List(Of Models.Attachment)) As String
            If String.IsNullOrEmpty(html) Then Return html
            If attachments Is Nothing OrElse attachments.Count = 0 Then Return html

            ' ContentId → 絶対ファイルパスの辞書を構築（大文字小文字を区別しない）
            Dim cidMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each att As Models.Attachment In attachments
                If att.IsInline AndAlso Not String.IsNullOrEmpty(att.ContentId) Then
                    If Not cidMap.ContainsKey(att.ContentId) Then
                        cidMap.Add(att.ContentId, att.FilePath)
                    End If
                End If
            Next

            If cidMap.Count = 0 Then Return html

            ' src="cid:xxx" または src='cid:xxx' を置換する
            Return Regex.Replace(
                html,
                "(?i)(src\s*=\s*[""'])cid:([^""']+)([""'])",
                Function(m As Match) As String
                    Dim cidVal As String = m.Groups(2).Value.Trim()
                    Dim filePath As String = Nothing
                    If cidMap.TryGetValue(cidVal, filePath) AndAlso File.Exists(filePath) Then
                        ' Windows パスを file:/// URL に変換（バックスラッシュをスラッシュに）
                        Dim fileUrl As String = "file:///" & filePath.Replace("\"c, "/"c)
                        Return m.Groups(1).Value & fileUrl & m.Groups(3).Value
                    End If
                    Return m.Value
                End Function)
        End Function

        ''' <summary>FTS クエリから最初の検索語を抽出する（"column:term" → "term"、引用符除去）。</summary>
        Private Function ExtractFirstTerm(query As String) As String
            Dim q As String = query.Trim()
            ' 列修飾子 "column:term" の列部分を除去
            Dim colonIdx As Integer = q.IndexOf(":"c)
            If colonIdx >= 0 AndAlso colonIdx < q.Length - 1 Then
                q = q.Substring(colonIdx + 1).Trim()
            End If
            ' 引用符を除去
            q = q.Replace("""", String.Empty).Trim()
            ' 最初のトークンのみ使用
            Dim spaceIdx As Integer = q.IndexOf(" "c)
            If spaceIdx > 0 Then q = q.Substring(0, spaceIdx)
            Return q.Trim()
        End Function

        ''' <summary>JavaScript 文字列リテラル用のエスケープ処理。</summary>
        Private Function JsEscape(s As String) As String
            Return s.Replace("\", "\\").Replace("'", "\'").Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
        End Function

        ''' <summary>
        ''' HTML 本文の &lt;/body&gt; 直前に検索語ハイライト用 JS を挿入して返す。
        ''' テキストノードを走査し、一致箇所を &lt;mark style="background:yellow"&gt; で囲む。
        ''' </summary>
        Private Function InjectHighlightScript(html As String, query As String) As String
            Dim term As String = ExtractFirstTerm(query)
            If String.IsNullOrEmpty(term) Then Return html

            Dim escaped As String = JsEscape(term)
            Dim script As String =
                "<script>" & vbCrLf &
                "(function() {" & vbCrLf &
                "  var q = '" & escaped & "'.toLowerCase();" & vbCrLf &
                "  if (!q) return;" & vbCrLf &
                "  function walk(node) {" & vbCrLf &
                "    if (node.nodeType === 3) {" & vbCrLf &
                "      var idx = node.data.toLowerCase().indexOf(q);" & vbCrLf &
                "      if (idx >= 0) {" & vbCrLf &
                "        var after = node.splitText(idx);" & vbCrLf &
                "        after.splitText(q.length);" & vbCrLf &
                "        var span = document.createElement('mark');" & vbCrLf &
                "        span.style.background = 'yellow';" & vbCrLf &
                "        span.style.color = 'black';" & vbCrLf &
                "        span.appendChild(after.cloneNode(true));" & vbCrLf &
                "        after.parentNode.replaceChild(span, after);" & vbCrLf &
                "      }" & vbCrLf &
                "    } else if (node.nodeType === 1 && node.tagName !== 'SCRIPT' && node.tagName !== 'STYLE') {" & vbCrLf &
                "      var kids = [];" & vbCrLf &
                "      for (var i = 0; i < node.childNodes.length; i++) { kids.push(node.childNodes[i]); }" & vbCrLf &
                "      for (var i = 0; i < kids.length; i++) { walk(kids[i]); }" & vbCrLf &
                "    }" & vbCrLf &
                "  }" & vbCrLf &
                "  if (document.body) walk(document.body);" & vbCrLf &
                "})();" & vbCrLf &
                "</script>"

            Dim bodyCloseIdx As Integer = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase)
            If bodyCloseIdx >= 0 Then
                Return html.Substring(0, bodyCloseIdx) & script & html.Substring(bodyCloseIdx)
            Else
                Return html & script
            End If
        End Function

        ''' <summary>JSON 配列文字列（["a@b.com","c@d.com"]）をカンマ区切りテキストに変換する。</summary>
        Private Function FormatRecipientsJson(json As String) As String
            If String.IsNullOrEmpty(json) Then Return String.Empty
            Dim s As String = json.Trim()
            If s.StartsWith("[") Then s = s.Substring(1)
            If s.EndsWith("]") Then s = s.Substring(0, s.Length - 1)
            Return s.Replace("""", String.Empty).Trim()
        End Function

        Private Sub LoadAttachments(attachments As List(Of Models.Attachment))
            flowAttachments.Controls.Clear()

            If attachments Is Nothing OrElse attachments.Count = 0 Then
                pnlAttachments.Visible = False
                Return
            End If

            Dim hasVisible As Boolean = False
            For Each att As Models.Attachment In attachments
                ' インライン画像は添付ファイル一覧に表示しない
                If att.IsInline Then Continue For

                Dim btn As New System.Windows.Forms.Button()
                btn.Text = att.FileName
                btn.Tag = att
                btn.Height = 26
                btn.AutoSize = True
                btn.Padding = New System.Windows.Forms.Padding(4, 0, 4, 0)
                AddHandler btn.Click, AddressOf AttachmentButton_Click
                flowAttachments.Controls.Add(btn)
                hasVisible = True
            Next
            pnlAttachments.Visible = hasVisible
        End Sub

        ' ════════════════════════════════════════════════════════════
        '  イベントハンドラ
        ' ════════════════════════════════════════════════════════════

        Private Sub btnToggleView_Click(sender As Object, e As EventArgs) Handles btnToggleView.Click
            If _currentEmail Is Nothing Then Return
            _showHtml = Not _showHtml
            If _showHtml Then
                Dim html As String = If(_currentEmail.BodyHtml, String.Empty)
                html = ReplaceCidReferences(html, _currentEmail.Attachments)
                If Not String.IsNullOrEmpty(_highlightQuery) Then
                    html = InjectHighlightScript(html, _highlightQuery)
                End If
                webBrowser.DocumentText = html
                webBrowser.Visible = True
                txtBodyText.Visible = False
                btnToggleView.Text = "テキスト表示"
            Else
                txtBodyText.Text = If(_currentEmail.BodyText, String.Empty)
                txtBodyText.Visible = True
                webBrowser.Visible = False
                btnToggleView.Text = "HTML 表示"
            End If
        End Sub

        Private Sub AttachmentButton_Click(sender As Object, e As EventArgs)
            Dim btn As System.Windows.Forms.Button = TryCast(sender, System.Windows.Forms.Button)
            If btn Is Nothing Then Return

            Dim att As Models.Attachment = TryCast(btn.Tag, Models.Attachment)
            If att Is Nothing Then Return

            If Not File.Exists(att.FilePath) Then
                MessageBox.Show("ファイルが見つかりません:" & vbCrLf & att.FilePath,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim ext As String = Path.GetExtension(att.FileName).ToLower()
            If ext = ".jpg" OrElse ext = ".jpeg" OrElse ext = ".png" OrElse
               ext = ".gif" OrElse ext = ".bmp" Then
                ShowImagePreview(att.FilePath, att.FileName)
            Else
                Try
                    Dim psi As New System.Diagnostics.ProcessStartInfo(att.FilePath)
                    psi.UseShellExecute = True
                    System.Diagnostics.Process.Start(psi)
                Catch ex As Exception
                    MessageBox.Show("ファイルを開けませんでした:" & vbCrLf & ex.Message,
                        "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Sub

        Private Sub ShowImagePreview(filePath As String, fileName As String)
            Dim img As System.Drawing.Image
            Try
                img = System.Drawing.Image.FromFile(filePath)
            Catch ex As Exception
                MessageBox.Show("画像を読み込めませんでした:" & vbCrLf & ex.Message,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Dim frm As New System.Windows.Forms.Form()
            frm.Text = fileName
            frm.Size = New System.Drawing.Size(800, 600)
            frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent

            Dim pb As New System.Windows.Forms.PictureBox()
            pb.Dock = System.Windows.Forms.DockStyle.Fill
            pb.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            pb.Image = img

            frm.Controls.Add(pb)
            frm.ShowDialog(Me)

            img.Dispose()
            frm.Dispose()
        End Sub

    End Class

End Namespace
