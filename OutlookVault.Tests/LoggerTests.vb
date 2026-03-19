Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookVault.Services
Imports System.IO

<TestFixture>
Public Class LoggerTests

    Private _testDir As String

    <SetUp>
    Public Sub SetUp()
        _testDir = Path.Combine(Path.GetTempPath(), "LoggerTests_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(_testDir)
        Logger.LogDirectory = _testDir
    End Sub

    <TearDown>
    Public Sub TearDown()
        Logger.LogDirectory = Nothing
        If Directory.Exists(_testDir) Then
            Directory.Delete(_testDir, True)
        End If
    End Sub

    <Test>
    Public Sub Info_WritesInfoLevelLog()
        Logger.Info("テストメッセージ")

        Dim logFile As String = Logger.GetLogFilePath()
        Assert.That(File.Exists(logFile), [Is].True)

        Dim content As String = File.ReadAllText(logFile)
        Assert.That(content, Does.Contain("[INFO]"))
        Assert.That(content, Does.Contain("テストメッセージ"))
    End Sub

    <Test>
    Public Sub Warn_WritesWarnLevelLog()
        Logger.Warn("警告メッセージ")

        Dim content As String = File.ReadAllText(Logger.GetLogFilePath())
        Assert.That(content, Does.Contain("[WARN]"))
        Assert.That(content, Does.Contain("警告メッセージ"))
    End Sub

    <Test>
    Public Sub Error_WritesErrorLevelLog()
        Logger.Error("エラーメッセージ")

        Dim content As String = File.ReadAllText(Logger.GetLogFilePath())
        Assert.That(content, Does.Contain("[ERROR]"))
        Assert.That(content, Does.Contain("エラーメッセージ"))
    End Sub

    <Test>
    Public Sub ErrorWithException_IncludesExceptionInfo()
        Dim ex As New InvalidOperationException("テスト例外")
        Logger.Error("処理失敗", ex)

        Dim content As String = File.ReadAllText(Logger.GetLogFilePath())
        Assert.That(content, Does.Contain("[ERROR]"))
        Assert.That(content, Does.Contain("処理失敗"))
        Assert.That(content, Does.Contain("InvalidOperationException"))
        Assert.That(content, Does.Contain("テスト例外"))
        ' スタックトレースが含まれていること
        Assert.That(content, Does.Contain("ErrorWithException_IncludesExceptionInfo"))
    End Sub

    <Test>
    Public Sub MultipleWrites_AppendsToSameFile()
        Logger.Info("1行目")
        Logger.Info("2行目")

        Dim lines() As String = File.ReadAllLines(Logger.GetLogFilePath())
        Assert.That(lines.Length, [Is].EqualTo(2))
    End Sub

    <Test>
    Public Sub LogLine_ContainsTimestamp()
        Logger.Info("タイムスタンプ確認")

        Dim content As String = File.ReadAllText(Logger.GetLogFilePath())
        Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
        Assert.That(content, Does.Contain(today))
    End Sub

    <Test>
    Public Sub GetLogFilePath_ContainsDateInFileName()
        Dim logFile As String = Logger.GetLogFilePath()
        Dim expectedDate As String = DateTime.Now.ToString("yyyyMMdd")
        Assert.That(Path.GetFileName(logFile), Does.Contain(expectedDate))
        Assert.That(Path.GetFileName(logFile), Does.StartWith("OutlookVault_"))
        Assert.That(Path.GetFileName(logFile), Does.EndWith(".log"))
    End Sub

End Class
