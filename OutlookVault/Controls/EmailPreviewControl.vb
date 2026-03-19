Option Explicit On
Option Strict On
Option Infer Off

Imports System.Drawing
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

        ''' <summary>受信者 JSON パース用の正規表現（コンパイル済みで再利用）。</summary>
        Private Shared ReadOnly RecipientJsonPattern As New Regex(
            "\{\s*""name""\s*:\s*""(?<name>[^""]*)""\s*,\s*""email""\s*:\s*""(?<email>[^""]*)""\s*\}",
            RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Private _currentEmail As Models.Email
        Private _showHtml As Boolean
        Private _canToggle As Boolean
        Private _highlightQuery As String
        ' _attachImageList, _attachLargeImageList, _attachContextMenu, _attachToolTip は Designer.vb で宣言・生成
        Private ReadOnly _htmlSanitizer As New Services.HtmlSanitizerService()

        ''' <summary>ダブルクリックで開くことをブロックする危険な拡張子。</summary>
        Private Shared ReadOnly BlockedExtensions() As String = {
            ".exe", ".bat", ".cmd", ".com", ".scr", ".pif",
            ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh", ".ws",
            ".ps1", ".ps2", ".psc1", ".psc2",
            ".msi", ".msp", ".mst",
            ".cpl", ".hta", ".inf", ".reg",
            ".dll", ".ocx", ".sys",
            ".lnk", ".url"
        }

        ' ════════════════════════════════════════════════════════════
        '  初期化
        ' ════════════════════════════════════════════════════════════

        Public Sub New()
            InitializeComponent()
            ' キャプションラベルにボールドフォントを適用
            Dim boldFont As New Font(Me.Font, FontStyle.Bold)
            lblFromCaption.Font = boldFont
            lblDateCaption.Font = boldFont
            lblSubjectCaption.Font = boldFont
            lblToCaption.Font = boldFont

            ' 添付ファイルアイコンを ImageList に登録
            BuildAttachmentIcons(_attachImageList, 16)
            BuildAttachmentIcons(_attachLargeImageList, 40)
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
            _canToggle = False
            _showHtml = False
        End Sub

        ''' <summary>HTML/テキスト表示の切り替えが可能かどうか。</summary>
        Public ReadOnly Property CanToggleView As Boolean
            Get
                Return _canToggle
            End Get
        End Property

        ''' <summary>現在 HTML 表示中かどうか。</summary>
        Public ReadOnly Property IsHtmlView As Boolean
            Get
                Return _showHtml
            End Get
        End Property

        ''' <summary>HTML/テキスト表示を切り替える。MainForm のボタンから呼ばれる。</summary>
        Public Sub ToggleView()
            If _currentEmail Is Nothing Then Return
            _showHtml = Not _showHtml
            If _showHtml Then
                Dim html As String = If(_currentEmail.BodyHtml, String.Empty)
                html = _htmlSanitizer.Sanitize(html)
                html = ReplaceCidReferences(html, _currentEmail.Attachments)
                If Not String.IsNullOrEmpty(_highlightQuery) Then
                    html = InjectHighlightScript(html, _highlightQuery)
                End If
                webBrowser.DocumentText = html
                webBrowser.Visible = True
                txtBodyText.Visible = False
            Else
                txtBodyText.Text = If(_currentEmail.BodyText, String.Empty)
                txtBodyText.Visible = True
                webBrowser.Visible = False
            End If
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
            _canToggle = hasHtml AndAlso hasText
            _showHtml = hasHtml

            If _showHtml Then
                Dim html As String = If(email.BodyHtml, String.Empty)
                html = _htmlSanitizer.Sanitize(html)
                html = ReplaceCidReferences(html, email.Attachments)
                If Not String.IsNullOrEmpty(_highlightQuery) Then
                    html = InjectHighlightScript(html, _highlightQuery)
                End If
                webBrowser.DocumentText = html
                webBrowser.Visible = True
                txtBodyText.Visible = False
            Else
                txtBodyText.Text = If(email.BodyText, String.Empty)
                txtBodyText.Visible = True
                webBrowser.Visible = False
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

        ''' <summary>検索クエリから最初の検索語を抽出する（"column:term" → "term"、引用符除去）。</summary>
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

        ''' <summary>
        ''' 受信者 JSON 配列文字列を「名前 &lt;メール&gt;」形式のカンマ区切りテキストに変換する。
        ''' 対応形式: [{"name":"太郎","email":"tarou@example.com"}, ...] または ["a@b.com", ...]
        ''' </summary>
        Private Function FormatRecipientsJson(json As String) As String
            If String.IsNullOrEmpty(json) Then Return String.Empty
            Dim s As String = json.Trim()
            If Not s.StartsWith("[") Then Return s

            ' {"name":"...","email":"..."} オブジェクト形式をパース
            Dim matches As MatchCollection = RecipientJsonPattern.Matches(s)

            If matches.Count > 0 Then
                Dim parts As New List(Of String)()
                For Each m As Match In matches
                    Dim recipName As String = m.Groups("name").Value.Trim()
                    Dim recipEmail As String = m.Groups("email").Value.Trim()
                    If recipName.Length > 0 Then
                        parts.Add(recipName)
                    ElseIf recipEmail.Length > 0 Then
                        parts.Add(recipEmail)
                    End If
                Next
                Return String.Join("; ", parts.ToArray())
            End If

            ' フォールバック: 単純な文字列配列 ["a@b.com","c@d.com"]
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

                ' 各添付ファイルを Panel(80×80) で表示
                Dim pnl As New System.Windows.Forms.Panel()
                pnl.Size = New Size(80, 80)
                pnl.Tag = att
                pnl.Cursor = System.Windows.Forms.Cursors.Hand
                pnl.ContextMenuStrip = _attachContextMenu

                ' アイコン (40×40) を中央上部に配置
                Dim pb As New System.Windows.Forms.PictureBox()
                pb.Size = New Size(40, 40)
                pb.Location = New Point(20, 4)
                pb.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
                Dim iconKey As String = GetIconKey(Path.GetExtension(att.FileName))
                If _attachLargeImageList.Images.ContainsKey(iconKey) Then
                    pb.Image = _attachLargeImageList.Images(iconKey)
                End If
                pb.Tag = att
                pb.Cursor = System.Windows.Forms.Cursors.Hand
                pb.ContextMenuStrip = _attachContextMenu

                ' ファイル名ラベル（省略表示、下部に配置）
                Dim lbl As New System.Windows.Forms.Label()
                lbl.Text = att.FileName
                lbl.Location = New Point(2, 48)
                lbl.Size = New Size(76, 30)
                lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter
                lbl.AutoEllipsis = True
                lbl.Tag = att
                lbl.Cursor = System.Windows.Forms.Cursors.Hand
                lbl.ContextMenuStrip = _attachContextMenu
                Dim smallFont As New Font(Me.Font.FontFamily, 7.5F, FontStyle.Regular)
                lbl.Font = smallFont

                ' ToolTip にフルファイル名とサイズを表示
                Dim tipText As String = att.FileName
                If att.FileSize > 0 Then
                    tipText = tipText & " (" & FormatFileSize(att.FileSize) & ")"
                End If
                _attachToolTip.SetToolTip(pnl, tipText)
                _attachToolTip.SetToolTip(pb, tipText)
                _attachToolTip.SetToolTip(lbl, tipText)

                ' ダブルクリックでファイルを開く（シングルクリックでは開かない）
                AddHandler pnl.DoubleClick, AddressOf AttachmentPanel_DoubleClick
                AddHandler pb.DoubleClick, AddressOf AttachmentPanel_DoubleClick
                AddHandler lbl.DoubleClick, AddressOf AttachmentPanel_DoubleClick

                ' ホバー時のハイライト（子コントロール間の移動でちらつかないよう遅延チェック）
                AddHandler pnl.MouseEnter, AddressOf AttachmentPanel_MouseEnter
                AddHandler pnl.MouseLeave, AddressOf AttachmentPanel_MouseLeave
                AddHandler pb.MouseEnter, AddressOf AttachmentChild_MouseEnter
                AddHandler pb.MouseLeave, AddressOf AttachmentChild_MouseLeave
                AddHandler lbl.MouseEnter, AddressOf AttachmentChild_MouseEnter
                AddHandler lbl.MouseLeave, AddressOf AttachmentChild_MouseLeave

                pnl.Controls.Add(pb)
                pnl.Controls.Add(lbl)
                flowAttachments.Controls.Add(pnl)
                hasVisible = True
            Next
            pnlAttachments.Visible = hasVisible
        End Sub

        ''' <summary>ファイルサイズを読みやすい形式に変換する。</summary>
        Private Shared Function FormatFileSize(bytes As Long) As String
            If bytes < 1024L Then
                Return bytes.ToString() & " B"
            ElseIf bytes < 1024L * 1024L Then
                Return (bytes / 1024.0).ToString("F1") & " KB"
            Else
                Return (bytes / (1024.0 * 1024.0)).ToString("F1") & " MB"
            End If
        End Function

        ' ════════════════════════════════════════════════════════════
        '  添付ファイルアイコン
        ' ════════════════════════════════════════════════════════════

        ''' <summary>ファイル種類ごとのアイコンを ImageList に登録する。</summary>
        Private Shared Sub BuildAttachmentIcons(imgList As System.Windows.Forms.ImageList, iconSize As Integer)
            ' PDF (赤)
            imgList.Images.Add("pdf", CreateFileIcon(Color.FromArgb(220, 50, 50), "P", iconSize))
            ' Word (青)
            imgList.Images.Add("doc", CreateFileIcon(Color.FromArgb(40, 90, 180), "W", iconSize))
            ' Excel (緑)
            imgList.Images.Add("xls", CreateFileIcon(Color.FromArgb(30, 130, 60), "X", iconSize))
            ' PowerPoint (オレンジ)
            imgList.Images.Add("ppt", CreateFileIcon(Color.FromArgb(210, 120, 20), "P", iconSize))
            ' 画像 (紫)
            imgList.Images.Add("img", CreateFileIcon(Color.FromArgb(120, 70, 170), "I", iconSize))
            ' テキスト (グレー)
            imgList.Images.Add("txt", CreateFileIcon(Color.FromArgb(100, 100, 100), "T", iconSize))
            ' 圧縮 (黄)
            imgList.Images.Add("zip", CreateFileIcon(Color.FromArgb(180, 150, 20), "Z", iconSize))
            ' その他 (グレー)
            imgList.Images.Add("other", CreateFileIcon(Color.FromArgb(140, 140, 140), "", iconSize))
        End Sub

        ''' <summary>指定サイズのファイルアイコンを生成する。</summary>
        Private Shared Function CreateFileIcon(baseColor As Color, letter As String, iconSize As Integer) As Bitmap
            Dim bmp As New Bitmap(iconSize, iconSize)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

                ' 角丸矩形（ファイル形状）
                Dim margin As Integer = CInt(Math.Max(1, iconSize * 0.06))
                Dim bodyW As Integer = iconSize - margin * 2
                Dim bodyH As Integer = iconSize - margin * 2
                Using brush As New SolidBrush(baseColor)
                    g.FillRectangle(brush, margin, margin, bodyW, bodyH)
                End Using
                ' 折り目（右上三角）
                Dim foldSize As Integer = CInt(iconSize * 0.3)
                Dim foldX As Integer = iconSize - margin - foldSize
                Dim pts() As Point = {
                    New Point(foldX, margin),
                    New Point(iconSize - margin, margin + foldSize),
                    New Point(iconSize - margin, margin)
                }
                Using brush As New SolidBrush(Color.FromArgb(60, Color.White))
                    g.FillPolygon(brush, pts)
                End Using

                ' 中央の文字
                If letter.Length > 0 Then
                    Dim fontSize As Single = CSng(iconSize * 0.45)
                    Using fnt As New Font("Segoe UI", fontSize, FontStyle.Bold)
                        Using sf As New StringFormat()
                            sf.Alignment = StringAlignment.Center
                            sf.LineAlignment = StringAlignment.Center
                            g.DrawString(letter, fnt, Brushes.White,
                                         New RectangleF(0, CSng(margin), CSng(iconSize), CSng(bodyH)), sf)
                        End Using
                    End Using
                End If
            End Using
            Return bmp
        End Function

        ''' <summary>拡張子からアイコンキーを返す。</summary>
        Private Shared Function GetIconKey(ext As String) As String
            If String.IsNullOrEmpty(ext) Then Return "other"
            Select Case ext.ToLower()
                Case ".pdf"
                    Return "pdf"
                Case ".doc", ".docx", ".docm", ".rtf"
                    Return "doc"
                Case ".xls", ".xlsx", ".xlsm", ".csv"
                    Return "xls"
                Case ".ppt", ".pptx", ".pptm"
                    Return "ppt"
                Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".svg", ".webp"
                    Return "img"
                Case ".txt", ".log", ".xml", ".json", ".html", ".htm", ".css", ".js", ".vb", ".cs"
                    Return "txt"
                Case ".zip", ".rar", ".7z", ".tar", ".gz", ".lzh"
                    Return "zip"
                Case Else
                    Return "other"
            End Select
        End Function

        ' ════════════════════════════════════════════════════════════
        '  イベントハンドラ
        ' ════════════════════════════════════════════════════════════

        Private Sub AttachmentSaveAs_Click(sender As Object, e As EventArgs) Handles menuSaveAs.Click
            ' コンテキストメニューの親コントロールから Attachment を取得
            Dim menuItem As System.Windows.Forms.ToolStripMenuItem =
                TryCast(sender, System.Windows.Forms.ToolStripMenuItem)
            If menuItem Is Nothing Then Return
            Dim cms As System.Windows.Forms.ContextMenuStrip =
                TryCast(menuItem.Owner, System.Windows.Forms.ContextMenuStrip)
            If cms Is Nothing Then Return
            Dim sourceCtl As System.Windows.Forms.Control = cms.SourceControl
            If sourceCtl Is Nothing Then Return

            Dim att As Models.Attachment = TryCast(sourceCtl.Tag, Models.Attachment)
            If att Is Nothing Then Return

            If Not File.Exists(att.FilePath) Then
                MessageBox.Show("ファイルが見つかりません:" & vbCrLf & att.FilePath,
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using dlg As New System.Windows.Forms.SaveFileDialog()
                dlg.FileName = att.FileName
                dlg.Filter = "すべてのファイル (*.*)|*.*"
                dlg.Title = "添付ファイルを保存"
                If dlg.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                    Try
                        File.Copy(att.FilePath, dlg.FileName, True)
                    Catch ex As Exception
                        MessageBox.Show("保存に失敗しました:" & vbCrLf & ex.Message,
                            "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Using
        End Sub

        Private Sub AttachmentPanel_DoubleClick(sender As Object, e As EventArgs)
            Dim ctl As System.Windows.Forms.Control = TryCast(sender, System.Windows.Forms.Control)
            If ctl Is Nothing Then Return

            Dim att As Models.Attachment = TryCast(ctl.Tag, Models.Attachment)
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
            ElseIf IsBlockedExtension(ext) Then
                MessageBox.Show(
                    "セキュリティ上の理由により、この種類のファイルは直接開けません。" & vbCrLf &
                    "右クリック→「名前を付けて保存」で保存してから開いてください。" & vbCrLf & vbCrLf &
                    "ファイル名: " & att.FileName,
                    "ブロックされたファイル", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

        Private Sub AttachmentPanel_MouseEnter(sender As Object, e As EventArgs)
            Dim pnl As System.Windows.Forms.Panel = TryCast(sender, System.Windows.Forms.Panel)
            If pnl IsNot Nothing Then pnl.BackColor = System.Drawing.SystemColors.ControlLight
        End Sub

        Private Sub AttachmentPanel_MouseLeave(sender As Object, e As EventArgs)
            Dim pnl As System.Windows.Forms.Panel = TryCast(sender, System.Windows.Forms.Panel)
            If pnl IsNot Nothing Then
                ' 子コントロールへの移動時にちらつかないよう遅延チェック
                pnl.BeginInvoke(DirectCast(Sub()
                    If Not pnl.IsDisposed Then
                        Dim cursorPos As Point = pnl.PointToClient(System.Windows.Forms.Cursor.Position)
                        If Not pnl.ClientRectangle.Contains(cursorPos) Then
                            pnl.BackColor = System.Drawing.Color.Transparent
                        End If
                    End If
                End Sub, MethodInvoker))
            End If
        End Sub

        Private Sub AttachmentChild_MouseEnter(sender As Object, e As EventArgs)
            Dim ctrl As System.Windows.Forms.Control = TryCast(sender, System.Windows.Forms.Control)
            If ctrl IsNot Nothing Then
                Dim pnl As System.Windows.Forms.Panel = TryCast(ctrl.Parent, System.Windows.Forms.Panel)
                If pnl IsNot Nothing Then pnl.BackColor = System.Drawing.SystemColors.ControlLight
            End If
        End Sub

        Private Sub AttachmentChild_MouseLeave(sender As Object, e As EventArgs)
            Dim ctrl As System.Windows.Forms.Control = TryCast(sender, System.Windows.Forms.Control)
            If ctrl IsNot Nothing Then
                Dim pnl As System.Windows.Forms.Panel = TryCast(ctrl.Parent, System.Windows.Forms.Panel)
                If pnl IsNot Nothing Then
                    pnl.BeginInvoke(DirectCast(Sub()
                        If Not pnl.IsDisposed Then
                            Dim cursorPos As Point = pnl.PointToClient(System.Windows.Forms.Cursor.Position)
                            If Not pnl.ClientRectangle.Contains(cursorPos) Then
                                pnl.BackColor = System.Drawing.Color.Transparent
                            End If
                        End If
                    End Sub, MethodInvoker))
                End If
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

        ''' <summary>指定された拡張子がブロック対象かどうかを返す。</summary>
        Private Shared Function IsBlockedExtension(ext As String) As Boolean
            If String.IsNullOrEmpty(ext) Then Return False
            For Each blocked As String In BlockedExtensions
                If String.Equals(ext, blocked, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' WebBrowser のナビゲーションを制御する。
        ''' about:blank（初期表示・DocumentText セット）以外のナビゲーションをブロックし、
        ''' 外部 URL の場合はデフォルトブラウザで開くか確認する。
        ''' </summary>
        Private Sub WebBrowser_Navigating(sender As Object, e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles webBrowser.Navigating
            Dim url As String = e.Url.ToString()

            ' about:blank は DocumentText の設定時に発生するため許可
            If url.Equals("about:blank", StringComparison.OrdinalIgnoreCase) Then Return

            ' それ以外のナビゲーションはキャンセル
            e.Cancel = True

            ' http/https リンクの場合、ユーザーに確認してデフォルトブラウザで開く
            If url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
               url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Dim result As System.Windows.Forms.DialogResult = MessageBox.Show(
                    "外部リンクをブラウザで開きますか？" & vbCrLf & vbCrLf & url,
                    "リンクを開く", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = System.Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim psi As New System.Diagnostics.ProcessStartInfo(url)
                        psi.UseShellExecute = True
                        System.Diagnostics.Process.Start(psi)
                    Catch ex As Exception
                        MessageBox.Show("リンクを開けませんでした:" & vbCrLf & ex.Message,
                            "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
        End Sub

    End Class

End Namespace
