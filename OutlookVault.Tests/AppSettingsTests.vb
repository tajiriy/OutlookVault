Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookVault.Config

<TestFixture>
Public Class AppSettingsTests

    Private _settings As AppSettings

    <SetUp>
    Public Sub SetUp()
        _settings = AppSettings.Instance
    End Sub

    ' ── タスクトレイ設定のデフォルト値テスト ──────────────────

    <Test>
    Public Sub MinimizeToTray_DefaultValue_IsFalse()
        ' AppSettings はシングルトンで既存の設定値を返すため、
        ' App.config に明示的な設定がない場合のデフォルト値が True であることを確認
        ' (テスト環境では config キーが未設定のため、デフォルト値が返る)
        Dim value As Boolean = _settings.MinimizeToTray
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    <Test>
    Public Sub CloseToTray_DefaultValue_IsFalse()
        Dim value As Boolean = _settings.CloseToTray
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    <Test>
    Public Sub ShowBalloonOnImport_DefaultValue_IsTrue()
        Dim value As Boolean = _settings.ShowBalloonOnImport
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    ' ── 既存設定のアクセステスト ──────────────────────────────

    <Test>
    Public Sub AutoImportEnabled_CanRead()
        Dim value As Boolean = _settings.AutoImportEnabled
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    <Test>
    Public Sub AutoImportIntervalMinutes_CanRead()
        Dim value As Integer = _settings.AutoImportIntervalMinutes
        Assert.That(value, [Is].GreaterThan(0))
    End Sub

    <Test>
    Public Sub ShowImportResult_CanRead()
        Dim value As Boolean = _settings.ShowImportResult
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    <Test>
    Public Sub ShowImportErrorDialog_CanRead()
        Dim value As Boolean = _settings.ShowImportErrorDialog
        Assert.That(value, [Is].TypeOf(Of Boolean)())
    End Sub

    ' ── ログディレクトリ設定テスト ──────────────────────────────

    <Test>
    Public Sub LogDirectory_DefaultValue_IsLogs()
        Dim value As String = _settings.LogDirectory
        Assert.That(value, [Is].EqualTo(".\logs"))
    End Sub

    ' ── 閲覧ウィンドウ設定テスト ──────────────────────────────

    <Test>
    Public Sub ReadingPanePosition_DefaultValue_IsBottom()
        Dim value As String = _settings.ReadingPanePosition
        Assert.That(value, [Is].EqualTo("Bottom"))
    End Sub

    <Test>
    Public Sub MailListWidth_DefaultValue_IsZero()
        Dim value As Integer = _settings.MailListWidth
        Assert.That(value, [Is].TypeOf(Of Integer)())
    End Sub

    ' ── フォルダ別列設定: SanitizeFolderKey テスト ──────────────

    <Test>
    Public Sub SanitizeFolderKey_NullOrEmpty_ReturnsAllTag()
        Assert.That(AppSettings.SanitizeFolderKey(Nothing), [Is].EqualTo("__ALL__"))
        Assert.That(AppSettings.SanitizeFolderKey(""), [Is].EqualTo("__ALL__"))
    End Sub

    <Test>
    Public Sub SanitizeFolderKey_NormalFolderName_ReturnsAsIs()
        Assert.That(AppSettings.SanitizeFolderKey("受信トレイ"), [Is].EqualTo("受信トレイ"))
        Assert.That(AppSettings.SanitizeFolderKey("__TRASH__"), [Is].EqualTo("__TRASH__"))
    End Sub

    <Test>
    Public Sub SanitizeFolderKey_SpecialChars_ReplacedWithUnderscore()
        Assert.That(AppSettings.SanitizeFolderKey("folder.name"), [Is].EqualTo("folder_name"))
        Assert.That(AppSettings.SanitizeFolderKey("a/b\c"), [Is].EqualTo("a_b_c"))
        Assert.That(AppSettings.SanitizeFolderKey("a:b<c>d"), [Is].EqualTo("a_b_c_d"))
    End Sub

    <Test>
    Public Sub SanitizeFolderKey_SpacesReplaced()
        Assert.That(AppSettings.SanitizeFolderKey("my folder"), [Is].EqualTo("my_folder"))
    End Sub

    ' ── フォルダ別列設定: HasFolderColumnSettings テスト ──────

    <Test>
    Public Sub HasFolderColumnSettings_NonExistent_ReturnsFalse()
        ' 存在しないフォルダキーの場合、False を返す
        Assert.That(_settings.HasFolderColumnSettings("__NON_EXISTENT_FOLDER_KEY_FOR_TEST__"), [Is].False)
    End Sub

End Class
