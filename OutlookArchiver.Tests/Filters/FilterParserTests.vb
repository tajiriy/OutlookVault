Option Explicit On
Option Strict On
Option Infer Off

Imports NUnit.Framework
Imports System.Data
Imports OutlookArchiver.Filters

<TestFixture>
Public Class FilterParserTests

    Private _table As DataTable

    <SetUp>
    Public Sub SetUp()
        _table = New DataTable()
        _table.Columns.Add("id", GetType(Integer))
        _table.Columns.Add("subject", GetType(String))
        _table.Columns.Add("sender_name", GetType(String))
        _table.Columns.Add("received_at", GetType(String))
        _table.Columns.Add("has_attachments", GetType(Integer))
    End Sub

    <TearDown>
    Public Sub TearDown()
        If _table IsNot Nothing Then _table.Dispose()
    End Sub

#Region "空・null 入力"

    <Test>
    Public Sub Parse_EmptyString_ReturnsEmpty()
        Assert.That(FilterParser.Parse("", _table.Columns), [Is].EqualTo(""))
    End Sub

    <Test>
    Public Sub Parse_WhitespaceOnly_ReturnsEmpty()
        Assert.That(FilterParser.Parse("   ", _table.Columns), [Is].EqualTo(""))
    End Sub

    <Test>
    Public Sub Parse_Nothing_ReturnsEmpty()
        Assert.That(FilterParser.Parse(Nothing, _table.Columns), [Is].EqualTo(""))
    End Sub

#End Region

#Region "単純テキスト（列指定なし）"

    <Test>
    Public Sub Parse_SimpleText_AllStringColumnsLike()
        Dim result As String = FilterParser.Parse("amazon", _table.Columns)
        ' 全文字列列（subject, sender_name, received_at）の OR
        Assert.That(result, Does.Contain("[subject] LIKE '%amazon%'"))
        Assert.That(result, Does.Contain("[sender_name] LIKE '%amazon%'"))
        Assert.That(result, Does.Contain("[received_at] LIKE '%amazon%'"))
        Assert.That(result, Does.Contain(" OR "))
    End Sub

#End Region

#Region "列指定: 部分一致"

    <Test>
    Public Sub Parse_ColumnColon_PartialMatch()
        Dim result As String = FilterParser.Parse("subject: amazon", _table.Columns)
        Assert.That(result, [Is].EqualTo("[subject] LIKE '%amazon%'"))
    End Sub

    <Test>
    Public Sub Parse_ColumnColon_CaseInsensitiveColumnName()
        Dim result As String = FilterParser.Parse("SUBJECT: amazon", _table.Columns)
        Assert.That(result, [Is].EqualTo("[subject] LIKE '%amazon%'"))
    End Sub

    <Test>
    Public Sub Parse_ColumnColon_NonStringColumn()
        Dim result As String = FilterParser.Parse("id: 1", _table.Columns)
        Assert.That(result, [Is].EqualTo("CONVERT([id], 'System.String') LIKE '%1%'"))
    End Sub

#End Region

#Region "列指定= 完全一致"

    <Test>
    Public Sub Parse_ColumnEquals_ExactMatch()
        Dim result As String = FilterParser.Parse("subject= amazon", _table.Columns)
        Assert.That(result, [Is].EqualTo("[subject] = 'amazon'"))
    End Sub

    <Test>
    Public Sub Parse_ColumnEquals_NonStringColumn()
        Dim result As String = FilterParser.Parse("id= 1", _table.Columns)
        Assert.That(result, [Is].EqualTo("CONVERT([id], 'System.String') = '1'"))
    End Sub

    <Test>
    Public Sub Parse_ColumnEquals_WithSpaces_ExactMatch()
        ' "id = 1" のように演算子の前後に空白がある場合
        Dim result As String = FilterParser.Parse("id = 1", _table.Columns)
        Assert.That(result, [Is].EqualTo("CONVERT([id], 'System.String') = '1'"))
    End Sub

    <Test>
    Public Sub Parse_ColumnColon_WithSpaceBefore_PartialMatch()
        ' "subject : amazon" のように演算子の前に空白がある場合
        Dim result As String = FilterParser.Parse("subject : amazon", _table.Columns)
        Assert.That(result, [Is].EqualTo("[subject] LIKE '%amazon%'"))
    End Sub

#End Region

#Region "AND 結合"

    <Test>
    Public Sub Parse_And_CombinesTwoConditions()
        Dim result As String = FilterParser.Parse("subject: amazon and sender_name: john", _table.Columns)
        Assert.That(result, Does.Contain("[subject] LIKE '%amazon%'"))
        Assert.That(result, Does.Contain("[sender_name] LIKE '%john%'"))
        Assert.That(result, Does.Contain(" AND "))
    End Sub

#End Region

#Region "OR 結合"

    <Test>
    Public Sub Parse_Or_CombinesTwoConditions()
        Dim result As String = FilterParser.Parse("subject: amazon or subject: google", _table.Columns)
        Assert.That(result, Does.Contain("[subject] LIKE '%amazon%'"))
        Assert.That(result, Does.Contain("[subject] LIKE '%google%'"))
        Assert.That(result, Does.Contain(" OR "))
    End Sub

#End Region

#Region "AND が OR より優先"

    <Test>
    Public Sub Parse_AndOrPrecedence_AndGroupedBeforeOr()
        ' a AND b OR c → (a AND b) OR c
        Dim result As String = FilterParser.Parse("subject: amazon and sender_name: john or received_at: 2025", _table.Columns)
        Assert.That(result, [Is].EqualTo("([subject] LIKE '%amazon%' AND [sender_name] LIKE '%john%') OR [received_at] LIKE '%2025%'"))
    End Sub

    <Test>
    Public Sub Parse_OrAndPrecedence_SecondAndGrouped()
        ' a OR b AND c → a OR (b AND c)
        Dim result As String = FilterParser.Parse("received_at: 2025 or subject: amazon and sender_name: john", _table.Columns)
        Assert.That(result, [Is].EqualTo("[received_at] LIKE '%2025%' OR ([subject] LIKE '%amazon%' AND [sender_name] LIKE '%john%')"))
    End Sub

    <Test>
    Public Sub Parse_MultipleAndOr_CorrectGrouping()
        ' a AND b OR c AND d → (a AND b) OR (c AND d)
        Dim result As String = FilterParser.Parse("subject: a and sender_name: b or subject: c and sender_name: d", _table.Columns)
        Assert.That(result, [Is].EqualTo("([subject] LIKE '%a%' AND [sender_name] LIKE '%b%') OR ([subject] LIKE '%c%' AND [sender_name] LIKE '%d%')"))
    End Sub

#End Region

#Region "存在しない列名"

    <Test>
    Public Sub Parse_NonExistentColumn_ConditionIgnored()
        Dim result As String = FilterParser.Parse("nonexistent: value", _table.Columns)
        Assert.That(result, [Is].EqualTo(""))
    End Sub

    <Test>
    Public Sub Parse_MixedExistentAndNonExistent_OnlyValidIncluded()
        Dim result As String = FilterParser.Parse("subject: amazon and nonexistent: value", _table.Columns)
        Assert.That(result, [Is].EqualTo("[subject] LIKE '%amazon%'"))
    End Sub

#End Region

#Region "特殊文字エスケープ"

    <Test>
    Public Sub Parse_SingleQuoteInValue_Escaped()
        Dim result As String = FilterParser.Parse("subject: it''s", _table.Columns)
        Assert.That(result, Does.Contain("it''''s"))
    End Sub

    <Test>
    Public Sub Parse_BracketInValue_Escaped()
        Dim result As String = FilterParser.Parse("subject: [test]", _table.Columns)
        Assert.That(result, Does.Contain("[[]test]"))
    End Sub

#End Region

#Region "大文字小文字無視（AND/OR キーワード）"

    <Test>
    Public Sub Parse_UppercaseAND_Recognized()
        Dim result As String = FilterParser.Parse("subject: a AND sender_name: b", _table.Columns)
        Assert.That(result, Does.Contain(" AND "))
    End Sub

    <Test>
    Public Sub Parse_MixedCaseOr_Recognized()
        Dim result As String = FilterParser.Parse("subject: a Or subject: b", _table.Columns)
        Assert.That(result, Does.Contain(" OR "))
    End Sub

#End Region

#Region "実データでの検証"

    <Test>
    Public Sub Parse_AppliedToDataTable_FiltersCorrectly()
        ' テストデータ投入
        _table.Rows.Add(1, "Amazon Order", "John", "2025-01-01", 0)
        _table.Rows.Add(2, "Google Alert", "Jane", "2025-02-01", 1)
        _table.Rows.Add(3, "Amazon Return", "Jane", "2025-03-01", 0)

        ' subject: amazon → 2件
        _table.DefaultView.RowFilter = FilterParser.Parse("subject: amazon", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(2))

        ' subject: amazon and sender_name: jane → 1件
        _table.DefaultView.RowFilter = FilterParser.Parse("subject: amazon and sender_name: jane", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(1))

        ' subject: google or sender_name: john → 2件
        _table.DefaultView.RowFilter = FilterParser.Parse("subject: google or sender_name: john", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(2))

        ' id= 2 → 1件（完全一致）
        _table.DefaultView.RowFilter = FilterParser.Parse("id= 2", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(1))
    End Sub

    <Test>
    Public Sub Parse_CaseInsensitiveValue_FiltersCorrectly()
        _table.Rows.Add(1, "Amazon Order", "John", "2025-01-01", 0)

        ' 大文字小文字を変えても同じ結果
        _table.DefaultView.RowFilter = FilterParser.Parse("subject: AMAZON", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(1))

        _table.DefaultView.RowFilter = FilterParser.Parse("subject: amazon", _table.Columns)
        Assert.That(_table.DefaultView.Count, [Is].EqualTo(1))
    End Sub

#End Region

End Class
