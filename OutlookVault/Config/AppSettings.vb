Option Explicit On
Option Strict On
Option Infer Off

Imports System.Configuration

Namespace Config

    ''' <summary>
    ''' アプリケーション設定の読み書きを担当するシングルトンクラス。
    ''' 値は App.config の appSettings セクションに永続化される。
    ''' </summary>
    Public Class AppSettings

        Private Shared ReadOnly _instance As New AppSettings()

        Public Shared ReadOnly Property Instance As AppSettings
            Get
                Return _instance
            End Get
        End Property

        ''' <summary>設定保存に失敗した場合に通知するイベント。引数はユーザー向けメッセージ。</summary>
        Public Event ConfigSaveError As EventHandler(Of String)

        Private Sub New()
        End Sub

        ' ── DB / ファイルパス ──────────────────────────────────────

        ''' <summary>ログファイルの出力ディレクトリ</summary>
        Public Property LogDirectory As String
            Get
                Dim val As String = ConfigurationManager.AppSettings("LogDirectory")
                Return If(String.IsNullOrEmpty(val), ".\logs", val)
            End Get
            Set(value As String)
                SaveSetting("LogDirectory", value)
            End Set
        End Property

        ''' <summary>SQLite データベースファイルのパス</summary>
        Public Property DbFilePath As String
            Get
                Dim val As String = ConfigurationManager.AppSettings("DbFilePath")
                Return If(String.IsNullOrEmpty(val), ".\data\archive.db", val)
            End Get
            Set(value As String)
                SaveSetting("DbFilePath", value)
            End Set
        End Property

        ''' <summary>添付ファイルの保存ディレクトリ</summary>
        Public Property AttachmentDirectory As String
            Get
                Dim val As String = ConfigurationManager.AppSettings("AttachmentDirectory")
                Return If(String.IsNullOrEmpty(val), ".\data\attachments\", val)
            End Get
            Set(value As String)
                SaveSetting("AttachmentDirectory", value)
            End Set
        End Property

        ' ── 自動取り込み ──────────────────────────────────────────

        ''' <summary>起動時に自動取り込みを開始するか</summary>
        Public Property AutoImportEnabled As Boolean
            Get
                Return GetBool("AutoImportEnabled", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("AutoImportEnabled", value.ToString())
            End Set
        End Property

        ''' <summary>自動取り込みモード（0: 間隔, 1: 定時）</summary>
        Public Property AutoImportMode As Integer
            Get
                Return GetInt("AutoImportMode", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("AutoImportMode", value.ToString())
            End Set
        End Property

        ''' <summary>自動取り込みのポーリング間隔（分）</summary>
        Public Property AutoImportIntervalMinutes As Integer
            Get
                Return GetInt("AutoImportIntervalMinutes", defaultValue:=10)
            End Get
            Set(value As Integer)
                SaveSetting("AutoImportIntervalMinutes", value.ToString())
            End Set
        End Property

        ''' <summary>定時取り込みの時刻（HH:mm 形式）</summary>
        Public Property ScheduledImportTime As String
            Get
                Dim val As String = ConfigurationManager.AppSettings("ScheduledImportTime")
                Return If(String.IsNullOrEmpty(val), "20:00", val)
            End Get
            Set(value As String)
                SaveSetting("ScheduledImportTime", value)
            End Set
        End Property

        ''' <summary>1回の取り込みで処理する最大件数</summary>
        Public Property MaxImportCount As Integer
            Get
                Return GetInt("MaxImportCount", defaultValue:=100)
            End Get
            Set(value As Integer)
                SaveSetting("MaxImportCount", value.ToString())
            End Set
        End Property

        ' ── 対象フォルダ ──────────────────────────────────────────

        ''' <summary>取り込み対象の Outlook フォルダ名一覧</summary>
        Public Property TargetFolders As List(Of String)
            Get
                Dim val As String = ConfigurationManager.AppSettings("TargetFolders")
                If String.IsNullOrEmpty(val) Then Return New List(Of String)({"受信トレイ", "送信済みアイテム"})
                Return New List(Of String)(val.Split(";"c))
            End Get
            Set(value As List(Of String))
                SaveSetting("TargetFolders", String.Join(";", value))
            End Set
        End Property

        ' ── 取り込み順序 ────────────────────────────────────────────

        ''' <summary>取り込み順序を古い順にするか（True: 古い順、False: 新しい順）</summary>
        Public Property ImportOldestFirst As Boolean
            Get
                Return GetBool("ImportOldestFirst", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("ImportOldestFirst", value.ToString())
            End Set
        End Property

        ' ── 同期モード ──────────────────────────────────────────────

        ''' <summary>同期モード（0: フルスキャン, 1: 差分スキャン）</summary>
        Public Property SyncMode As Integer
            Get
                Return GetInt("SyncMode", defaultValue:=1)
            End Get
            Set(value As Integer)
                SaveSetting("SyncMode", value.ToString())
            End Set
        End Property

        ''' <summary>差分スキャン時のバッファ時間（時間単位）</summary>
        Public Property DiffSyncBufferHours As Integer
            Get
                Return GetInt("DiffSyncBufferHours", defaultValue:=24)
            End Get
            Set(value As Integer)
                SaveSetting("DiffSyncBufferHours", value.ToString())
            End Set
        End Property

        ' ── 削除同期 ────────────────────────────────────────────────

        ''' <summary>取り込み時に Outlook 側で削除されたメールをデータベースからも削除するか</summary>
        Public Property SyncDeletionsEnabled As Boolean
            Get
                Return GetBool("SyncDeletionsEnabled", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("SyncDeletionsEnabled", value.ToString())
            End Set
        End Property

        ' ── タスクトレイ常駐 ──────────────────────────────────────

        ''' <summary>最小化時にタスクトレイに格納するか</summary>
        Public Property MinimizeToTray As Boolean
            Get
                Return GetBool("MinimizeToTray", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("MinimizeToTray", value.ToString())
            End Set
        End Property

        ''' <summary>閉じるボタンでタスクトレイに格納するか</summary>
        Public Property CloseToTray As Boolean
            Get
                Return GetBool("CloseToTray", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("CloseToTray", value.ToString())
            End Set
        End Property

        ''' <summary>Windows 起動時にアプリを自動起動するか（レジストリ Run キーで管理）</summary>
        Public Property StartWithWindows As Boolean
            Get
                Try
                    Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        "Software\Microsoft\Windows\CurrentVersion\Run", writable:=False)
                        If key Is Nothing Then Return False
                        Dim val As Object = key.GetValue("OutlookVault")
                        Return val IsNot Nothing
                    End Using
                Catch ex As Exception
                    ' レジストリアクセス失敗（権限不足等）は False を返す
                    Return False
                End Try
            End Get
            Set(value As Boolean)
                Try
                    Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        "Software\Microsoft\Windows\CurrentVersion\Run", writable:=True)
                        If key Is Nothing Then Return
                        If value Then
                            Dim exePath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                            key.SetValue("OutlookVault", """" & exePath & """ --minimized")
                        Else
                            key.DeleteValue("OutlookVault", throwOnMissingValue:=False)
                        End If
                    End Using
                Catch ex As Security.SecurityException
                    Services.Logger.Warn("Windows スタートアップ登録に失敗しました（権限不足）: " & ex.Message)
                Catch ex As UnauthorizedAccessException
                    Services.Logger.Warn("Windows スタートアップ登録に失敗しました（アクセス拒否）: " & ex.Message)
                End Try
            End Set
        End Property

        ''' <summary>取り込み完了時にバルーン通知を表示するか</summary>
        Public Property ShowBalloonOnImport As Boolean
            Get
                Return GetBool("ShowBalloonOnImport", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("ShowBalloonOnImport", value.ToString())
            End Set
        End Property

        ' ── ゴミ箱 ────────────────────────────────────────────────

        ''' <summary>ゴミ箱の自動削除日数（0 で自動削除無効、デフォルト: 30日）</summary>
        Public Property TrashAutoDeleteDays As Integer
            Get
                Return GetInt("TrashAutoDeleteDays", defaultValue:=30)
            End Get
            Set(value As Integer)
                SaveSetting("TrashAutoDeleteDays", value.ToString())
            End Set
        End Property

        ' ── 結果ダイアログ ────────────────────────────────────────

        ''' <summary>取り込み完了時に結果ダイアログを表示するか</summary>
        Public Property ShowImportResult As Boolean
            Get
                Return GetBool("ShowImportResult", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("ShowImportResult", value.ToString())
            End Set
        End Property

        ''' <summary>取り込みエラー時に結果ダイアログを表示するか</summary>
        Public Property ShowImportErrorDialog As Boolean
            Get
                Return GetBool("ShowImportErrorDialog", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("ShowImportErrorDialog", value.ToString())
            End Set
        End Property

        ' ── 表示設定 ──────────────────────────────────────────────

        ''' <summary>会話ビューで古い順（昇順）に表示するか</summary>
        Public Property ConversationSortAscending As Boolean
            Get
                Return GetBool("ConversationSortAscending", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("ConversationSortAscending", value.ToString())
            End Set
        End Property

        ''' <summary>デフォルトで HTML 表示するか（False の場合はプレーンテキスト）</summary>
        Public Property DefaultHtmlView As Boolean
            Get
                Return GetBool("DefaultHtmlView", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("DefaultHtmlView", value.ToString())
            End Set
        End Property

        ' ── レイアウト ──────────────────────────────────────────────

        ''' <summary>フォルダツリーの幅（SplitContainer の SplitterDistance）</summary>
        Public Property FolderTreeWidth As Integer
            Get
                Return GetInt("FolderTreeWidth", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("FolderTreeWidth", value.ToString())
            End Set
        End Property

        ''' <summary>会話ビューのスレッド一覧の高さ（SplitContainer の SplitterDistance）</summary>
        Public Property ConversationSplitterDistance As Integer
            Get
                Return GetInt("ConversationSplitterDistance", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("ConversationSplitterDistance", value.ToString())
            End Set
        End Property

        ''' <summary>閲覧ウィンドウの位置（Bottom / Right / Off）</summary>
        Public Property ReadingPanePosition As String
            Get
                Dim val As String = ConfigurationManager.AppSettings("ReadingPanePosition")
                Return If(String.IsNullOrEmpty(val), "Bottom", val)
            End Get
            Set(value As String)
                SaveSetting("ReadingPanePosition", value)
            End Set
        End Property

        ''' <summary>メール一覧の高さ（splitRight の SplitterDistance、閲覧ウィンドウ「下」用）</summary>
        Public Property MailListHeight As Integer
            Get
                Return GetInt("MailListHeight", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("MailListHeight", value.ToString())
            End Set
        End Property

        ''' <summary>メール一覧の幅（splitRight の SplitterDistance、閲覧ウィンドウ「右」用）</summary>
        Public Property MailListWidth As Integer
            Get
                Return GetInt("MailListWidth", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("MailListWidth", value.ToString())
            End Set
        End Property

        ''' <summary>ウィンドウの幅</summary>
        Public Property WindowWidth As Integer
            Get
                Return GetInt("WindowWidth", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("WindowWidth", value.ToString())
            End Set
        End Property

        ''' <summary>ウィンドウの高さ</summary>
        Public Property WindowHeight As Integer
            Get
                Return GetInt("WindowHeight", defaultValue:=0)
            End Get
            Set(value As Integer)
                SaveSetting("WindowHeight", value.ToString())
            End Set
        End Property

        ''' <summary>ウィンドウの X 座標</summary>
        Public Property WindowLeft As Integer
            Get
                Return GetInt("WindowLeft", defaultValue:=Integer.MinValue)
            End Get
            Set(value As Integer)
                SaveSetting("WindowLeft", value.ToString())
            End Set
        End Property

        ''' <summary>ウィンドウの Y 座標</summary>
        Public Property WindowTop As Integer
            Get
                Return GetInt("WindowTop", defaultValue:=Integer.MinValue)
            End Get
            Set(value As Integer)
                SaveSetting("WindowTop", value.ToString())
            End Set
        End Property

        ''' <summary>ウィンドウが最大化されていたか</summary>
        Public Property WindowMaximized As Boolean
            Get
                Return GetBool("WindowMaximized", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("WindowMaximized", value.ToString())
            End Set
        End Property

        ' ── メール一覧表示設定 ────────────────────────────────────────

        ''' <summary>メール一覧の列幅（列インデックス順にピクセル値をカンマ区切りで格納）</summary>
        Public Property EmailListColumnWidths As String
            Get
                Return ConfigurationManager.AppSettings("EmailListColumnWidths")
            End Get
            Set(value As String)
                SaveSetting("EmailListColumnWidths", value)
            End Set
        End Property

        ''' <summary>メール一覧の列表示順（列インデックス順に DisplayIndex をカンマ区切りで格納）</summary>
        Public Property EmailListColumnOrder As String
            Get
                Return ConfigurationManager.AppSettings("EmailListColumnOrder")
            End Get
            Set(value As String)
                SaveSetting("EmailListColumnOrder", value)
            End Set
        End Property

        ''' <summary>メール一覧のソート列インデックス（0=添付 1=件名 2=差出人 3=受信日時 4=サイズ）</summary>
        Public Property EmailListSortColumn As Integer
            Get
                Return GetInt("EmailListSortColumn", defaultValue:=CInt(EmailListColumn.ReceivedAt))
            End Get
            Set(value As Integer)
                SaveSetting("EmailListSortColumn", value.ToString())
            End Set
        End Property

        ''' <summary>メール一覧のソート方向（True: 昇順、False: 降順）</summary>
        Public Property EmailListSortAscending As Boolean
            Get
                Return GetBool("EmailListSortAscending", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("EmailListSortAscending", value.ToString())
            End Set
        End Property

        ' ── フォルダ別メール一覧表示設定 ─────────────────────────

        ''' <summary>フォルダ別の列設定を保存する。</summary>
        ''' <param name="folderKey">フォルダ識別キー（フォルダ名 or "__ALL__" or "__TRASH__"）</param>
        ''' <param name="widths">列幅（カンマ区切り）</param>
        ''' <param name="order">列順（カンマ区切り）</param>
        ''' <param name="sortColumn">ソート列インデックス</param>
        ''' <param name="sortAscending">ソート昇順か</param>
        Public Sub SaveFolderColumnSettings(folderKey As String, widths As String, order As String, sortColumn As Integer, sortAscending As Boolean)
            Dim safeKey As String = SanitizeFolderKey(folderKey)
            SaveSetting("FolderCol_" & safeKey & "_Widths", widths)
            SaveSetting("FolderCol_" & safeKey & "_Order", order)
            SaveSetting("FolderCol_" & safeKey & "_SortCol", sortColumn.ToString())
            SaveSetting("FolderCol_" & safeKey & "_SortAsc", sortAscending.ToString())
        End Sub

        ''' <summary>フォルダ別の列設定を読み込む。未設定の場合は Nothing を返す。</summary>
        Public Function LoadFolderColumnWidths(folderKey As String) As String
            Return ConfigurationManager.AppSettings("FolderCol_" & SanitizeFolderKey(folderKey) & "_Widths")
        End Function

        ''' <summary>フォルダ別の列順を読み込む。未設定の場合は Nothing を返す。</summary>
        Public Function LoadFolderColumnOrder(folderKey As String) As String
            Return ConfigurationManager.AppSettings("FolderCol_" & SanitizeFolderKey(folderKey) & "_Order")
        End Function

        ''' <summary>フォルダ別のソート列を読み込む。未設定の場合は -1 を返す。</summary>
        Public Function LoadFolderSortColumn(folderKey As String) As Integer
            Return GetInt("FolderCol_" & SanitizeFolderKey(folderKey) & "_SortCol", defaultValue:=-1)
        End Function

        ''' <summary>フォルダ別のソート方向を読み込む。未設定の場合は False を返す。</summary>
        Public Function LoadFolderSortAscending(folderKey As String) As Boolean
            Return GetBool("FolderCol_" & SanitizeFolderKey(folderKey) & "_SortAsc", defaultValue:=False)
        End Function

        ''' <summary>フォルダ別の列設定が存在するか確認する。</summary>
        Public Function HasFolderColumnSettings(folderKey As String) As Boolean
            Dim val As String = ConfigurationManager.AppSettings("FolderCol_" & SanitizeFolderKey(folderKey) & "_Widths")
            Return Not String.IsNullOrEmpty(val)
        End Function

        ''' <summary>フォルダ名をApp.configキーとして安全な文字列に変換する。</summary>
        Friend Shared Function SanitizeFolderKey(folderKey As String) As String
            If String.IsNullOrEmpty(folderKey) Then Return "__ALL__"
            ' App.config キーで使用不可な文字をアンダースコアに置換
            Dim sanitized As String = folderKey
            For Each c As Char In New Char() {"."c, " "c, "\"c, "/"c, ":"c, "<"c, ">"c, """"c, "|"c, "?"c, "*"c}
                sanitized = sanitized.Replace(c, "_"c)
            Next
            Return sanitized
        End Function

        ' ── ヘルパー ──────────────────────────────────────────────

        Private Function GetBool(key As String, defaultValue As Boolean) As Boolean
            Dim val As String = ConfigurationManager.AppSettings(key)
            If String.IsNullOrEmpty(val) Then Return defaultValue
            Dim result As Boolean
            If Boolean.TryParse(val, result) Then Return result
            Return defaultValue
        End Function

        Private Function GetInt(key As String, defaultValue As Integer) As Integer
            Dim val As String = ConfigurationManager.AppSettings(key)
            Dim result As Integer
            If Integer.TryParse(val, result) Then Return result
            Return defaultValue
        End Function

        ''' <summary>設定保存エラーのメッセージボックスを一度だけ表示するためのフラグ</summary>
        Private Shared _configErrorShown As Boolean = False

        Private Sub SaveSetting(key As String, value As String)
            Dim maxRetries As Integer = 2
            For attempt As Integer = 1 To maxRetries
                Try
                    Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                    If config.AppSettings.Settings(key) Is Nothing Then
                        config.AppSettings.Settings.Add(key, value)
                    Else
                        config.AppSettings.Settings(key).Value = value
                    End If
                    config.Save(ConfigurationSaveMode.Modified)
                    ConfigurationManager.RefreshSection("appSettings")
                    Return
                Catch ex As ConfigurationErrorsException
                    ConfigurationManager.RefreshSection("appSettings")
                    If attempt = maxRetries Then
                        Services.Logger.Warn("設定の保存に失敗しました（構成ファイルが外部から変更されています）: " &
                                             key & " = " & value & " — " & ex.Message)
                        If Not _configErrorShown Then
                            _configErrorShown = True
                            RaiseEvent ConfigSaveError(Me,
                                "設定の保存に失敗しました。" & Environment.NewLine &
                                "構成ファイルが別のプログラム（アンチウィルスソフト等）によって変更されています。" & Environment.NewLine &
                                Environment.NewLine &
                                "設定は次回起動時に反映されない場合があります。")
                        End If
                    End If
                End Try
            Next
        End Sub

    End Class

End Namespace
