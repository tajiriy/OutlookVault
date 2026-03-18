Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Services

<TestFixture>
Public Class ImportLogWriterTests

    Private _tempDir As String

    <SetUp>
    Public Sub SetUp()
        _tempDir = IO.Path.Combine(IO.Path.GetTempPath(), "ImportLogWriterTests_" & Guid.NewGuid().ToString("N"))
        IO.Directory.CreateDirectory(_tempDir)
    End Sub

    <TearDown>
    Public Sub TearDown()
        If IO.Directory.Exists(_tempDir) Then
            IO.Directory.Delete(_tempDir, True)
        End If
    End Sub

    ' ── SummarizeErrors ─────────────────────────────────────

    <Test>
    Public Sub SummarizeErrors_GroupsByErrorMessage()
        Dim errors As New List(Of ImportErrorEntry)()
        errors.Add(New ImportErrorEntry("受信トレイ", "msg1", "件名1", "アクセスが拒否されました"))
        errors.Add(New ImportErrorEntry("受信トレイ", "msg2", "件名2", "タイムアウトしました"))
        errors.Add(New ImportErrorEntry("受信トレイ", "msg3", "件名3", "アクセスが拒否されました"))

        Dim summary As List(Of KeyValuePair(Of String, Integer)) = ImportLogWriter.SummarizeErrors(errors)

        Assert.AreEqual(2, summary.Count)
        ' 件数の多い順にソート
        Assert.AreEqual("アクセスが拒否されました", summary(0).Key)
        Assert.AreEqual(2, summary(0).Value)
        Assert.AreEqual("タイムアウトしました", summary(1).Key)
        Assert.AreEqual(1, summary(1).Value)
    End Sub

    <Test>
    Public Sub SummarizeErrors_EmptyList_ReturnsEmpty()
        Dim errors As New List(Of ImportErrorEntry)()

        Dim summary As List(Of KeyValuePair(Of String, Integer)) = ImportLogWriter.SummarizeErrors(errors)

        Assert.AreEqual(0, summary.Count)
    End Sub

    <Test>
    Public Sub SummarizeErrors_SingleErrorType_ReturnsSingleEntry()
        Dim errors As New List(Of ImportErrorEntry)()
        errors.Add(New ImportErrorEntry("受信トレイ", "msg1", "件名1", "エラーA"))
        errors.Add(New ImportErrorEntry("受信トレイ", "msg2", "件名2", "エラーA"))

        Dim summary As List(Of KeyValuePair(Of String, Integer)) = ImportLogWriter.SummarizeErrors(errors)

        Assert.AreEqual(1, summary.Count)
        Assert.AreEqual("エラーA", summary(0).Key)
        Assert.AreEqual(2, summary(0).Value)
    End Sub

    ' ── WriteErrorLog ───────────────────────────────────────

    <Test>
    Public Sub WriteErrorLog_EmptyList_ReturnsNothing()
        Dim errors As New List(Of ImportErrorEntry)()

        Dim result As String = ImportLogWriter.WriteErrorLog(errors)

        Assert.IsNull(result)
    End Sub

    <Test>
    Public Sub WriteErrorLog_NothingList_ReturnsNothing()
        Dim result As String = ImportLogWriter.WriteErrorLog(Nothing)

        Assert.IsNull(result)
    End Sub

    ' ── ImportErrorEntry ────────────────────────────────────

    <Test>
    Public Sub ImportErrorEntry_SetsTimestamp()
        Dim before As DateTime = DateTime.Now
        Dim entry As New ImportErrorEntry("フォルダ", "msg-id", "件名", "エラー")
        Dim after As DateTime = DateTime.Now

        Assert.GreaterOrEqual(entry.Timestamp, before)
        Assert.LessOrEqual(entry.Timestamp, after)
    End Sub

    <Test>
    Public Sub ImportErrorEntry_SetsAllProperties()
        Dim entry As New ImportErrorEntry("受信トレイ", "msg-123", "テスト件名", "テストエラー")

        Assert.AreEqual("受信トレイ", entry.FolderName)
        Assert.AreEqual("msg-123", entry.MessageId)
        Assert.AreEqual("テスト件名", entry.Subject)
        Assert.AreEqual("テストエラー", entry.ErrorMessage)
    End Sub

    ' ── ImportResult ────────────────────────────────────────

    <Test>
    Public Sub ImportResult_ErrorsInitializedAsEmptyList()
        Dim result As New ImportResult()

        Assert.IsNotNull(result.Errors)
        Assert.AreEqual(0, result.Errors.Count)
    End Sub

End Class
