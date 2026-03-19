Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookVault.Services

<TestFixture>
Public Class ImportServiceTests

    ' ── IsSentFolder テスト ──────────────────────────────

    <TestCase("送信済みアイテム", True)>
    <TestCase("送信トレイ", True)>
    <TestCase("Sent Items", True)>
    <TestCase("Sent Mail", True)>
    <TestCase("Sent Messages", True)>
    <TestCase("Outbox", True)>
    <TestCase("sent items", True)>
    <TestCase("SENT ITEMS", True)>
    <TestCase("受信トレイ", False)>
    <TestCase("Inbox", False)>
    <TestCase("下書き", False)>
    <TestCase("Drafts", False)>
    <TestCase("", False)>
    Public Sub IsSentFolder_ReturnsExpectedResult(folderName As String, expected As Boolean)
        Assert.That(ImportService.IsSentFolder(folderName), [Is].EqualTo(expected))
    End Sub

    <Test>
    Public Sub IsSentFolder_NullInput_ReturnsFalse()
        Assert.That(ImportService.IsSentFolder(Nothing), [Is].False)
    End Sub

End Class
