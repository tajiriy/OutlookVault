Option Explicit On
Option Strict On
Option Infer Off

Imports System.Data.SQLite
Imports System.Text

Namespace Filters

    ''' <summary>
    ''' メール検索用フィルタ。FilterParser のトークナイザ/パーサーを再利用し、
    ''' ユーザー入力を SQLite WHERE 句 + パラメータに変換する。
    '''
    ''' 日本語列名とDB列名の両方を受け付ける。
    '''   件名: amazon           → e.subject LIKE '%amazon%'
    '''   添付: あり             → e.has_attachments = 1
    '''   件名: amazon and 添付: あり → AND 結合
    ''' </summary>
    Public Class EmailSearchFilter

        ''' <summary>SQL WHERE 句とパラメータを保持する結果クラス。</summary>
        Public Class SearchQuery
            ''' <summary>WHERE 句（"WHERE " は含まない）。空の場合はフィルタなし。</summary>
            Public Property WhereClause As String = ""
            ''' <summary>パラメータ一覧。</summary>
            Public Property Parameters As New List(Of SQLiteParameter)()
        End Class

        ''' <summary>列マッピング定義。</summary>
        Private Class ColumnMapping
            Public Property DbColumn As String       ' DB列名 (e.subject 等)
            Public Property IsBoolean As Boolean     ' あり/なし 判定列か
            Public Property IsAttachmentTable As Boolean  ' attachments テーブル参照か
        End Class

        ''' <summary>日本語名・英語名 → DB列のマッピング。</summary>
        Private Shared ReadOnly ColumnMap As New Dictionary(Of String, ColumnMapping)(StringComparer.OrdinalIgnoreCase) From {
            {"件名", New ColumnMapping() With {.DbColumn = "e.subject"}},
            {"subject", New ColumnMapping() With {.DbColumn = "e.subject"}},
            {"本文", New ColumnMapping() With {.DbColumn = "e.body_text"}},
            {"body_text", New ColumnMapping() With {.DbColumn = "e.body_text"}},
            {"差出人", New ColumnMapping() With {.DbColumn = "e.sender_name"}},
            {"sender_name", New ColumnMapping() With {.DbColumn = "e.sender_name"}},
            {"メール", New ColumnMapping() With {.DbColumn = "e.sender_email"}},
            {"sender_email", New ColumnMapping() With {.DbColumn = "e.sender_email"}},
            {"添付ファイル", New ColumnMapping() With {.DbColumn = "a.file_name", .IsAttachmentTable = True}},
            {"file_name", New ColumnMapping() With {.DbColumn = "a.file_name", .IsAttachmentTable = True}},
            {"添付", New ColumnMapping() With {.DbColumn = "e.has_attachments", .IsBoolean = True}},
            {"has_attachments", New ColumnMapping() With {.DbColumn = "e.has_attachments", .IsBoolean = True}},
            {"フォルダ", New ColumnMapping() With {.DbColumn = "e.folder_name"}},
            {"folder_name", New ColumnMapping() With {.DbColumn = "e.folder_name"}}
        }

        ''' <summary>あり/なし → 1/0 の変換マップ。</summary>
        Private Shared ReadOnly BooleanTrueValues() As String = {"あり", "true", "1", "yes"}
        Private Shared ReadOnly BooleanFalseValues() As String = {"なし", "false", "0", "no"}

        Private _paramIndex As Integer = 0

        ''' <summary>
        ''' フィルタ文字列を SQLite WHERE 句に変換する。
        ''' </summary>
        ''' <param name="input">ユーザー入力のフィルタ文字列</param>
        ''' <param name="folderName">フォルダフィルタ（Nothing の場合は無視）</param>
        ''' <returns>SearchQuery。フィルタなしの場合は WhereClause が空。</returns>
        Public Function Parse(input As String, Optional folderName As String = Nothing) As SearchQuery
            Dim result As New SearchQuery()
            _paramIndex = 0

            If String.IsNullOrWhiteSpace(input) Then
                ' フィルタなし — フォルダ指定のみ
                If Not String.IsNullOrEmpty(folderName) Then
                    result.WhereClause = "e.folder_name = @folder"
                    result.Parameters.Add(New SQLiteParameter("@folder", folderName))
                End If
                Return result
            End If

            ' FilterParser のトークナイザ・パーサーを再利用
            Dim tokens As List(Of FilterParser.Token) = FilterParser.Tokenize(input.Trim())
            If tokens.Count = 0 Then Return result

            Dim conditions As List(Of FilterParser.Condition) = FilterParser.ParseConditions(tokens)
            If conditions.Count = 0 Then Return result

            ' OR で区切って AND グループに分割
            Dim orGroups As List(Of List(Of FilterParser.Condition)) = SplitByOr(conditions)

            ' SQL WHERE 句を構築
            Dim whereClause As String = BuildWhereClause(orGroups, result.Parameters)

            ' フォルダフィルタを追加
            If Not String.IsNullOrEmpty(folderName) Then
                Dim folderParam As String = NextParam()
                result.Parameters.Add(New SQLiteParameter(folderParam, folderName))
                If String.IsNullOrEmpty(whereClause) Then
                    whereClause = "e.folder_name = " & folderParam
                Else
                    whereClause = "(" & whereClause & ") AND e.folder_name = " & folderParam
                End If
            End If

            result.WhereClause = whereClause
            Return result
        End Function

        ''' <summary>条件リストを OR で区切って AND グループに分割する。</summary>
        Private Shared Function SplitByOr(conditions As List(Of FilterParser.Condition)) As List(Of List(Of FilterParser.Condition))
            Dim groups As New List(Of List(Of FilterParser.Condition))()
            Dim current As New List(Of FilterParser.Condition)()

            For Each cond As FilterParser.Condition In conditions
                current.Add(cond)
                If String.Equals(cond.FollowingOperator, "OR", StringComparison.OrdinalIgnoreCase) Then
                    groups.Add(current)
                    current = New List(Of FilterParser.Condition)()
                End If
            Next

            If current.Count > 0 Then
                groups.Add(current)
            End If

            Return groups
        End Function

        ''' <summary>OR グループ群から WHERE 句を生成する。</summary>
        Private Function BuildWhereClause(orGroups As List(Of List(Of FilterParser.Condition)),
                                          params As List(Of SQLiteParameter)) As String
            Dim orParts As New List(Of String)()

            For Each group As List(Of FilterParser.Condition) In orGroups
                Dim andParts As New List(Of String)()
                For Each cond As FilterParser.Condition In group
                    Dim expr As String = BuildConditionSql(cond, params)
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

        ''' <summary>1つの条件を SQL 式に変換する。</summary>
        Private Function BuildConditionSql(cond As FilterParser.Condition,
                                           params As List(Of SQLiteParameter)) As String
            If String.IsNullOrEmpty(cond.ColumnName) Then
                ' 列指定なし → 全検索対象列を OR で部分一致
                Return BuildAllColumnsLike(cond.Value, params)
            End If

            ' 列名マッピングを検索
            Dim mapping As ColumnMapping = Nothing
            If Not ColumnMap.TryGetValue(cond.ColumnName, mapping) Then
                Return "" ' 未知の列名は無視
            End If

            ' あり/なし 判定列
            If mapping.IsBoolean Then
                Return BuildBooleanCondition(mapping.DbColumn, cond.Value)
            End If

            ' 添付テーブル参照
            If mapping.IsAttachmentTable Then
                Return BuildAttachmentCondition(cond, params)
            End If

            ' 通常の列
            Dim paramName As String = NextParam()
            If cond.IsExact Then
                params.Add(New SQLiteParameter(paramName, cond.Value))
                Return String.Format("{0} = {1} COLLATE NOCASE", mapping.DbColumn, paramName)
            Else
                params.Add(New SQLiteParameter(paramName, "%" & cond.Value & "%"))
                Return String.Format("{0} LIKE {1}", mapping.DbColumn, paramName)
            End If
        End Function

        ''' <summary>全検索対象列を OR で部分一致する SQL を生成する。</summary>
        Private Function BuildAllColumnsLike(value As String, params As List(Of SQLiteParameter)) As String
            Dim likeValue As String = "%" & value & "%"
            Dim parts As New List(Of String)()

            ' メール本体の検索対象列
            Dim mainColumns() As String = {"e.subject", "e.body_text", "e.sender_name", "e.sender_email"}
            For Each col As String In mainColumns
                Dim p As String = NextParam()
                params.Add(New SQLiteParameter(p, likeValue))
                parts.Add(String.Format("{0} LIKE {1}", col, p))
            Next

            ' 添付ファイル名
            Dim attachParam As String = NextParam()
            params.Add(New SQLiteParameter(attachParam, likeValue))
            parts.Add(String.Format("e.id IN (SELECT a.email_id FROM attachments a WHERE a.file_name LIKE {0})", attachParam))

            Return "(" & String.Join(" OR ", parts) & ")"
        End Function

        ''' <summary>あり/なし 判定の SQL を生成する。</summary>
        Private Shared Function BuildBooleanCondition(dbColumn As String, value As String) As String
            If Array.Exists(BooleanTrueValues, Function(v) String.Equals(v, value, StringComparison.OrdinalIgnoreCase)) Then
                Return String.Format("{0} = 1", dbColumn)
            ElseIf Array.Exists(BooleanFalseValues, Function(v) String.Equals(v, value, StringComparison.OrdinalIgnoreCase)) Then
                Return String.Format("{0} = 0", dbColumn)
            End If
            Return "" ' 不明な値は無視
        End Function

        ''' <summary>添付ファイルテーブル参照の SQL を生成する。</summary>
        Private Function BuildAttachmentCondition(cond As FilterParser.Condition,
                                                  params As List(Of SQLiteParameter)) As String
            Dim paramName As String = NextParam()
            If cond.IsExact Then
                params.Add(New SQLiteParameter(paramName, cond.Value))
                Return String.Format("e.id IN (SELECT a.email_id FROM attachments a WHERE a.file_name = {0} COLLATE NOCASE)", paramName)
            Else
                params.Add(New SQLiteParameter(paramName, "%" & cond.Value & "%"))
                Return String.Format("e.id IN (SELECT a.email_id FROM attachments a WHERE a.file_name LIKE {0})", paramName)
            End If
        End Function

        ''' <summary>ユニークなパラメータ名を生成する。</summary>
        Private Function NextParam() As String
            _paramIndex += 1
            Return "@p" & _paramIndex.ToString()
        End Function

    End Class

End Namespace
