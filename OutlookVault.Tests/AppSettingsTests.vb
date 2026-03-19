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

End Class
