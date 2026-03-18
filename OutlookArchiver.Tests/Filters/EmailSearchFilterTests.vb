Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports OutlookArchiver.Filters

<TestFixture>
Public Class EmailSearchFilterTests

    Private _filter As EmailSearchFilter

    <SetUp>
    Public Sub SetUp()
        _filter = New EmailSearchFilter()
    End Sub

#Region "空・null 入力"

    <Test>
    Public Sub Parse_EmptyString_ReturnsEmptyWhereClause()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("")
        Assert.That(result.WhereClause, [Is].EqualTo(""))
        Assert.That(result.Parameters.Count, [Is].EqualTo(0))
    End Sub

    <Test>
    Public Sub Parse_Nothing_ReturnsEmptyWhereClause()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse(Nothing)
        Assert.That(result.WhereClause, [Is].EqualTo(""))
    End Sub

#End Region

#Region "単純テキスト（列指定なし）"

    <Test>
    Public Sub Parse_SimpleText_AllColumnsLike()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("amazon")
        ' 全検索対象列の OR
        Assert.That(result.WhereClause, Does.Contain("e.subject LIKE"))
        Assert.That(result.WhereClause, Does.Contain("e.body_text LIKE"))
        Assert.That(result.WhereClause, Does.Contain("e.sender_name LIKE"))
        Assert.That(result.WhereClause, Does.Contain("e.sender_email LIKE"))
        Assert.That(result.WhereClause, Does.Contain("attachments"))
        Assert.That(result.Parameters.Count, [Is].EqualTo(5))
        ' パラメータの値は %amazon%
        Assert.That(DirectCast(result.Parameters(0).Value, String), [Is].EqualTo("%amazon%"))
    End Sub

#End Region

#Region "日本語列名: 部分一致"

    <Test>
    Public Sub Parse_JapaneseColumnName_Subject()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名: amazon")
        Assert.That(result.WhereClause, [Is].EqualTo("e.subject LIKE @p1"))
        Assert.That(DirectCast(result.Parameters(0).Value, String), [Is].EqualTo("%amazon%"))
    End Sub

    <Test>
    Public Sub Parse_JapaneseColumnName_Body()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("本文: テスト")
        Assert.That(result.WhereClause, [Is].EqualTo("e.body_text LIKE @p1"))
        Assert.That(DirectCast(result.Parameters(0).Value, String), [Is].EqualTo("%テスト%"))
    End Sub

    <Test>
    Public Sub Parse_JapaneseColumnName_Sender()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("差出人: 田中")
        Assert.That(result.WhereClause, [Is].EqualTo("e.sender_name LIKE @p1"))
    End Sub

#End Region

#Region "英語列名"

    <Test>
    Public Sub Parse_EnglishColumnName_Subject()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("subject: amazon")
        Assert.That(result.WhereClause, [Is].EqualTo("e.subject LIKE @p1"))
    End Sub

    <Test>
    Public Sub Parse_EnglishColumnName_CaseInsensitive()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("SUBJECT: amazon")
        Assert.That(result.WhereClause, [Is].EqualTo("e.subject LIKE @p1"))
    End Sub

#End Region

#Region "完全一致（= 演算子）"

    <Test>
    Public Sub Parse_ExactMatch_UsesEquals()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名= テスト件名")
        Assert.That(result.WhereClause, Does.Contain("e.subject = @p1 COLLATE NOCASE"))
        Assert.That(DirectCast(result.Parameters(0).Value, String), [Is].EqualTo("テスト件名"))
    End Sub

#End Region

#Region "添付有無（あり/なし）"

    <Test>
    Public Sub Parse_AttachmentAri_HasAttachments1()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("添付: あり")
        Assert.That(result.WhereClause, [Is].EqualTo("e.has_attachments = 1"))
        Assert.That(result.Parameters.Count, [Is].EqualTo(0))
    End Sub

    <Test>
    Public Sub Parse_AttachmentNashi_HasAttachments0()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("添付: なし")
        Assert.That(result.WhereClause, [Is].EqualTo("e.has_attachments = 0"))
    End Sub

    <Test>
    Public Sub Parse_HasAttachments_True()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("has_attachments: true")
        Assert.That(result.WhereClause, [Is].EqualTo("e.has_attachments = 1"))
    End Sub

    <Test>
    Public Sub Parse_HasAttachments_False()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("has_attachments: false")
        Assert.That(result.WhereClause, [Is].EqualTo("e.has_attachments = 0"))
    End Sub

#End Region

#Region "添付ファイル名検索"

    <Test>
    Public Sub Parse_AttachmentFileName_SubQuery()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("添付ファイル: report")
        Assert.That(result.WhereClause, Does.Contain("attachments"))
        Assert.That(result.WhereClause, Does.Contain("file_name LIKE"))
        Assert.That(DirectCast(result.Parameters(0).Value, String), [Is].EqualTo("%report%"))
    End Sub

#End Region

#Region "AND 結合"

    <Test>
    Public Sub Parse_And_CombinesTwoConditions()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名: amazon and 添付: あり")
        Assert.That(result.WhereClause, Does.Contain("e.subject LIKE"))
        Assert.That(result.WhereClause, Does.Contain("e.has_attachments = 1"))
        Assert.That(result.WhereClause, Does.Contain(" AND "))
    End Sub

#End Region

#Region "OR 結合"

    <Test>
    Public Sub Parse_Or_CombinesTwoConditions()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名: amazon or 件名: google")
        Assert.That(result.WhereClause, Does.Contain(" OR "))
    End Sub

#End Region

#Region "AND/OR 優先順位"

    <Test>
    Public Sub Parse_AndOrPrecedence_AndGroupedBeforeOr()
        ' a AND b OR c → (a AND b) OR c
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名: amazon and 差出人: john or 本文: test")
        Assert.That(result.WhereClause, Does.StartWith("(e.subject LIKE"))
        Assert.That(result.WhereClause, Does.Contain(" AND "))
        Assert.That(result.WhereClause, Does.Contain(") OR "))
    End Sub

#End Region

#Region "フォルダフィルタ"

    <Test>
    Public Sub Parse_WithFolderName_AddsFolderCondition()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("件名: amazon", "受信トレイ")
        Assert.That(result.WhereClause, Does.Contain("e.folder_name ="))
        Assert.That(result.WhereClause, Does.Contain("e.subject LIKE"))
    End Sub

    <Test>
    Public Sub Parse_EmptyQueryWithFolder_FolderOnly()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("", "受信トレイ")
        Assert.That(result.WhereClause, [Is].EqualTo("e.folder_name = @folder"))
        Assert.That(result.Parameters.Count, [Is].EqualTo(1))
    End Sub

#End Region

#Region "存在しない列名"

    <Test>
    Public Sub Parse_NonExistentColumn_Ignored()
        Dim result As EmailSearchFilter.SearchQuery = _filter.Parse("存在しない列: value")
        Assert.That(result.WhereClause, [Is].EqualTo(""))
    End Sub

#End Region

End Class
