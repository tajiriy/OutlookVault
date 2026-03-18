Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data
Imports System.Text

Namespace Filters

    ''' <summary>
    ''' テーブルビューア用のフィルタ文字列をパースし、DataTable.RowFilter 式に変換する。
    '''
    ''' 構文:
    '''   テキスト           → 全文字列列を部分一致 (OR)
    '''   列名: 値           → 指定列で部分一致 (LIKE)
    '''   列名= 値           → 指定列で完全一致 (=)
    '''   条件 and 条件      → AND 結合
    '''   条件 or 条件       → OR 結合
    '''   AND は OR より優先  → a AND b OR c → (a AND b) OR c
    '''
    ''' 列名・値ともに大文字小文字を無視する。
    ''' </summary>
    Public Class FilterParser

        ''' <summary>
        ''' フィルタ入力文字列を DataTable.DefaultView.RowFilter 式に変換する。
        ''' </summary>
        ''' <param name="input">ユーザー入力のフィルタ文字列</param>
        ''' <param name="columns">対象テーブルの DataColumnCollection</param>
        ''' <returns>RowFilter 式文字列。フィルタなしの場合は空文字列。</returns>
        Public Shared Function Parse(input As String, columns As DataColumnCollection) As String
            If String.IsNullOrWhiteSpace(input) Then Return ""

            Dim trimmed As String = input.Trim()

            ' トークン列を生成
            Dim tokens As List(Of Token) = Tokenize(trimmed)
            If tokens.Count = 0 Then Return ""

            ' 条件リストにパース
            Dim conditions As List(Of Condition) = ParseConditions(tokens)
            If conditions.Count = 0 Then Return ""

            ' OR で区切って AND グループに分割
            Dim orGroups As List(Of List(Of Condition)) = SplitByOr(conditions)

            ' RowFilter 式を構築
            Return BuildRowFilter(orGroups, columns)
        End Function

#Region "Token"

        Friend Enum TokenType
            ColumnOperator  ' "列名:" または "列名="
            Value           ' 値テキスト
            [And]           ' "and" キーワード
            [Or]            ' "or" キーワード
        End Enum

        Friend Structure Token
            Public Type As TokenType
            Public Text As String          ' Value の場合は値テキスト
            Public ColumnName As String    ' ColumnOperator の場合は列名
            Public IsExact As Boolean      ' ColumnOperator の場合: True=完全一致, False=部分一致
        End Structure

        ''' <summary>入力文字列をトークン列に分解する。</summary>
        Friend Shared Function Tokenize(input As String) As List(Of Token)
            Dim tokens As New List(Of Token)()
            Dim i As Integer = 0
            Dim length As Integer = input.Length

            While i < length
                ' 空白をスキップ
                If Char.IsWhiteSpace(input(i)) Then
                    i += 1
                    Continue While
                End If

                ' 列名: 値 または 列名= 値 のパターンを試行（キーワードより先にチェック）
                ' "subject : amazon" のように列名の後に空白がある場合も認識するため
                Dim colToken As Token = Nothing
                Dim consumed As Integer = 0
                If TryMatchColumnOperator(input, i, colToken, consumed) Then
                    tokens.Add(colToken)
                    i += consumed
                    Continue While
                End If

                ' "and" / "or" キーワードのチェック（前後が区切り位置であること）
                Dim keyword As TokenType
                Dim keywordLen As Integer = 0
                If TryMatchKeyword(input, i, keyword, keywordLen) Then
                    tokens.Add(New Token() With {.Type = keyword, .Text = ""})
                    i += keywordLen
                    Continue While
                End If

                ' それ以外は値トークンとして読み取る（次の空白または末尾まで）
                Dim valueStart As Integer = i
                While i < length AndAlso Not Char.IsWhiteSpace(input(i))
                    i += 1
                End While
                tokens.Add(New Token() With {
                    .Type = TokenType.Value,
                    .Text = input.Substring(valueStart, i - valueStart)
                })
            End While

            Return tokens
        End Function

        ''' <summary>
        ''' 現在位置から "and" または "or" キーワードにマッチするか試行する。
        ''' キーワードの前後が単語境界であることを確認する。
        ''' </summary>
        Private Shared Function TryMatchKeyword(input As String, pos As Integer, ByRef keyword As TokenType, ByRef keywordLen As Integer) As Boolean
            Dim remaining As Integer = input.Length - pos

            ' "and" チェック
            If remaining >= 3 Then
                Dim candidate As String = input.Substring(pos, 3)
                If String.Equals(candidate, "and", StringComparison.OrdinalIgnoreCase) Then
                    If IsWordBoundary(input, pos, 3) Then
                        keyword = TokenType.And
                        keywordLen = 3
                        Return True
                    End If
                End If
            End If

            ' "or" チェック
            If remaining >= 2 Then
                Dim candidate As String = input.Substring(pos, 2)
                If String.Equals(candidate, "or", StringComparison.OrdinalIgnoreCase) Then
                    If IsWordBoundary(input, pos, 2) Then
                        keyword = TokenType.Or
                        keywordLen = 2
                        Return True
                    End If
                End If
            End If

            Return False
        End Function

        ''' <summary>指定位置の単語が前後の単語境界で区切られているか。</summary>
        Private Shared Function IsWordBoundary(input As String, pos As Integer, length As Integer) As Boolean
            ' 前方チェック: 先頭 or 直前が空白
            If pos > 0 AndAlso Not Char.IsWhiteSpace(input(pos - 1)) Then Return False
            ' 後方チェック: 末尾 or 直後が空白
            Dim endPos As Integer = pos + length
            If endPos < input.Length AndAlso Not Char.IsWhiteSpace(input(endPos)) Then Return False
            Return True
        End Function

        ''' <summary>
        ''' 現在位置から "列名:" または "列名=" パターンにマッチするか試行する。
        ''' </summary>
        Private Shared Function TryMatchColumnOperator(input As String, pos As Integer, ByRef token As Token, ByRef consumed As Integer) As Boolean
            Dim length As Integer = input.Length

            ' 列名部分を読み取る（英数字・アンダースコア）
            Dim nameStart As Integer = pos
            Dim i As Integer = pos
            While i < length AndAlso (Char.IsLetterOrDigit(input(i)) OrElse input(i) = "_"c)
                i += 1
            End While

            ' 列名が空ならマッチしない
            If i = nameStart Then Return False

            Dim nameEnd As Integer = i

            ' 列名と演算子の間の空白をスキップ（"id = 1" 対応）
            While i < length AndAlso Char.IsWhiteSpace(input(i))
                i += 1
            End While

            ' 演算子チェック（: または =）
            If i >= length Then Return False
            Dim op As Char = input(i)
            If op <> ":"c AndAlso op <> "="c Then Return False

            Dim colName As String = input.Substring(nameStart, nameEnd - nameStart)
            Dim isExact As Boolean = (op = "="c)

            ' 演算子の直後の空白をスキップ
            i += 1
            While i < length AndAlso Char.IsWhiteSpace(input(i))
                i += 1
            End While

            token = New Token() With {
                .Type = TokenType.ColumnOperator,
                .ColumnName = colName,
                .IsExact = isExact,
                .Text = ""
            }
            consumed = i - pos
            Return True
        End Function

#End Region

#Region "Condition"

        ''' <summary>パース済みの1つのフィルタ条件。</summary>
        Friend Structure Condition ' Friend: テストからアクセス可能
            ''' <summary>列名。空の場合は全列対象。</summary>
            Public ColumnName As String
            ''' <summary>True=完全一致, False=部分一致。</summary>
            Public IsExact As Boolean
            ''' <summary>フィルタ値。</summary>
            Public Value As String
            ''' <summary>直後の論理演算子。Nothing=なし, "AND", "OR"。</summary>
            Public FollowingOperator As String
        End Structure

        ''' <summary>トークン列を Condition リストに変換する。</summary>
        Friend Shared Function ParseConditions(tokens As List(Of Token)) As List(Of Condition)
            Dim result As New List(Of Condition)()
            Dim i As Integer = 0

            While i < tokens.Count
                Dim t As Token = tokens(i)

                Select Case t.Type
                    Case TokenType.ColumnOperator
                        ' 次のトークンが Value であればそれを値とする
                        Dim value As String = ""
                        If i + 1 < tokens.Count AndAlso tokens(i + 1).Type = TokenType.Value Then
                            value = tokens(i + 1).Text
                            i += 2
                        Else
                            ' 値なし — 条件をスキップ
                            i += 1
                            Continue While
                        End If

                        Dim cond As New Condition()
                        cond.ColumnName = t.ColumnName
                        cond.IsExact = t.IsExact
                        cond.Value = value
                        cond.FollowingOperator = Nothing

                        ' 後続の AND/OR を確認
                        If i < tokens.Count Then
                            If tokens(i).Type = TokenType.And Then
                                cond.FollowingOperator = "AND"
                                i += 1
                            ElseIf tokens(i).Type = TokenType.Or Then
                                cond.FollowingOperator = "OR"
                                i += 1
                            End If
                        End If

                        result.Add(cond)

                    Case TokenType.Value
                        ' 列指定なしの自由テキスト
                        Dim cond As New Condition()
                        cond.ColumnName = ""
                        cond.IsExact = False
                        cond.Value = t.Text
                        cond.FollowingOperator = Nothing
                        i += 1

                        ' 後続の AND/OR を確認
                        If i < tokens.Count Then
                            If tokens(i).Type = TokenType.And Then
                                cond.FollowingOperator = "AND"
                                i += 1
                            ElseIf tokens(i).Type = TokenType.Or Then
                                cond.FollowingOperator = "OR"
                                i += 1
                            End If
                        End If

                        result.Add(cond)

                    Case TokenType.And, TokenType.Or
                        ' 先頭や連続した演算子はスキップ
                        i += 1

                End Select
            End While

            Return result
        End Function

        ''' <summary>条件リストを OR で区切って AND グループに分割する。</summary>
        Private Shared Function SplitByOr(conditions As List(Of Condition)) As List(Of List(Of Condition))
            Dim groups As New List(Of List(Of Condition))()
            Dim current As New List(Of Condition)()

            For Each cond As Condition In conditions
                current.Add(cond)
                If String.Equals(cond.FollowingOperator, "OR", StringComparison.OrdinalIgnoreCase) Then
                    groups.Add(current)
                    current = New List(Of Condition)()
                End If
            Next

            If current.Count > 0 Then
                groups.Add(current)
            End If

            Return groups
        End Function

#End Region

#Region "RowFilter 生成"

        ''' <summary>OR グループ群から RowFilter 式を生成する。</summary>
        Private Shared Function BuildRowFilter(orGroups As List(Of List(Of Condition)), columns As DataColumnCollection) As String
            Dim orParts As New List(Of String)()

            For Each group As List(Of Condition) In orGroups
                Dim andParts As New List(Of String)()
                For Each cond As Condition In group
                    Dim expr As String = BuildConditionExpression(cond, columns)
                    If Not String.IsNullOrEmpty(expr) Then
                        andParts.Add(expr)
                    End If
                Next

                If andParts.Count = 1 Then
                    orParts.Add(andParts(0))
                ElseIf andParts.Count > 1 Then
                    orParts.Add("(" & String.Join(" AND ", andParts) & ")")
                End If
            Next

            If orParts.Count = 0 Then Return ""
            Return String.Join(" OR ", orParts)
        End Function

        ''' <summary>1つの条件を RowFilter 式に変換する。</summary>
        Private Shared Function BuildConditionExpression(cond As Condition, columns As DataColumnCollection) As String
            Dim escaped As String = EscapeFilterValue(cond.Value)

            If String.IsNullOrEmpty(cond.ColumnName) Then
                ' 列指定なし → 全文字列列を OR で部分一致
                Return BuildAllColumnsLike(escaped, columns)
            End If

            ' 列名を大文字小文字無視で検索
            Dim col As DataColumn = FindColumn(cond.ColumnName, columns)
            If col Is Nothing Then Return ""

            If cond.IsExact Then
                ' 完全一致（= 演算子）
                If col.DataType Is GetType(String) Then
                    Return String.Format("[{0}] = '{1}'", col.ColumnName, escaped)
                Else
                    Return String.Format("CONVERT([{0}], 'System.String') = '{1}'", col.ColumnName, escaped)
                End If
            Else
                ' 部分一致（: 演算子）
                If col.DataType Is GetType(String) Then
                    Return String.Format("[{0}] LIKE '%{1}%'", col.ColumnName, escaped)
                Else
                    Return String.Format("CONVERT([{0}], 'System.String') LIKE '%{1}%'", col.ColumnName, escaped)
                End If
            End If
        End Function

        ''' <summary>全文字列列を対象に LIKE 部分一致の OR 式を生成する。</summary>
        Private Shared Function BuildAllColumnsLike(escapedValue As String, columns As DataColumnCollection) As String
            Dim parts As New List(Of String)()

            For Each col As DataColumn In columns
                If col.DataType Is GetType(String) Then
                    parts.Add(String.Format("[{0}] LIKE '%{1}%'", col.ColumnName, escapedValue))
                End If
            Next

            ' 文字列列がない場合は CONVERT で対応
            If parts.Count = 0 Then
                For Each col As DataColumn In columns
                    parts.Add(String.Format("CONVERT([{0}], 'System.String') LIKE '%{1}%'", col.ColumnName, escapedValue))
                Next
            End If

            If parts.Count = 0 Then Return ""
            If parts.Count = 1 Then Return parts(0)
            Return "(" & String.Join(" OR ", parts) & ")"
        End Function

        ''' <summary>列名を大文字小文字無視で検索する。</summary>
        Private Shared Function FindColumn(name As String, columns As DataColumnCollection) As DataColumn
            For Each col As DataColumn In columns
                If String.Equals(col.ColumnName, name, StringComparison.OrdinalIgnoreCase) Then
                    Return col
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>RowFilter 式用に値をエスケープする。</summary>
        Private Shared Function EscapeFilterValue(value As String) As String
            Return value.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
        End Function

#End Region

    End Class

End Namespace
