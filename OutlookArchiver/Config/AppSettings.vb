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

        Private Sub New()
        End Sub

        ' ── DB / ファイルパス ──────────────────────────────────────

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
                Return GetBool("AutoImportEnabled", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("AutoImportEnabled", value.ToString())
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

        ''' <summary>アーカイブ対象の Outlook フォルダ名一覧</summary>
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

        ' ── 削除同期 ────────────────────────────────────────────────

        ''' <summary>取り込み時に Outlook 側で削除されたメールをアーカイブからも削除するか</summary>
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
                Return GetBool("MinimizeToTray", defaultValue:=True)
            End Get
            Set(value As Boolean)
                SaveSetting("MinimizeToTray", value.ToString())
            End Set
        End Property

        ''' <summary>閉じるボタンでタスクトレイに格納するか</summary>
        Public Property CloseToTray As Boolean
            Get
                Return GetBool("CloseToTray", defaultValue:=True)
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
                        Dim val As Object = key.GetValue("OutlookArchiver")
                        Return val IsNot Nothing
                    End Using
                Catch
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
                            key.SetValue("OutlookArchiver", """" & exePath & """ --minimized")
                        Else
                            key.DeleteValue("OutlookArchiver", throwOnMissingValue:=False)
                        End If
                    End Using
                Catch
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

        ' ── 結果ダイアログ ────────────────────────────────────────

        ''' <summary>取り込み完了時に結果ダイアログを表示するか</summary>
        Public Property ShowImportResult As Boolean
            Get
                Return GetBool("ShowImportResult", defaultValue:=False)
            End Get
            Set(value As Boolean)
                SaveSetting("ShowImportResult", value.ToString())
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
                Return GetInt("EmailListSortColumn", defaultValue:=3)
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

        Private Sub SaveSetting(key As String, value As String)
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            If config.AppSettings.Settings(key) Is Nothing Then
                config.AppSettings.Settings.Add(key, value)
            Else
                config.AppSettings.Settings(key).Value = value
            End If
            config.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection("appSettings")
        End Sub

    End Class

End Namespace
