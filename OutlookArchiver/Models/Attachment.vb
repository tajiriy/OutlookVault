Option Explicit On
Option Strict On
Option Infer Off

Namespace Models

    Public Class Attachment

        Public Property Id As Integer
        Public Property EmailId As Integer
        Public Property FileName As String
        ''' <summary>添付ファイルの保存先パス（attachments/ 配下の相対パス）</summary>
        Public Property FilePath As String
        Public Property FileSize As Long
        Public Property MimeType As String
        Public Property CreatedAt As DateTime
        ''' <summary>MIME Content-ID（インライン画像の cid: 参照に対応）</summary>
        Public Property ContentId As String
        ''' <summary>True の場合はメール本文中にインライン表示される画像（添付ファイル一覧には表示しない）</summary>
        Public Property IsInline As Boolean

    End Class

End Namespace
