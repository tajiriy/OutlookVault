Option Explicit On
Option Strict On
Option Infer Off

Namespace Models

    Public Class Email

        Public Property Id As Integer
        Public Property MessageId As String
        Public Property InReplyTo As String
        Public Property References As String
        Public Property ThreadId As String
        Public Property EntryId As String
        Public Property Subject As String
        Public Property NormalizedSubject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        ''' <summary>宛先リスト（JSON 配列文字列）</summary>
        Public Property ToRecipients As String
        ''' <summary>CC リスト（JSON 配列文字列）</summary>
        Public Property CcRecipients As String
        ''' <summary>BCC リスト（JSON 配列文字列）</summary>
        Public Property BccRecipients As String
        Public Property BodyText As String
        Public Property BodyHtml As String
        Public Property ReceivedAt As DateTime
        Public Property SentAt As Nullable(Of DateTime)
        Public Property FolderName As String
        Public Property HasAttachments As Boolean
        Public Property EmailSize As Long
        Public Property CreatedAt As DateTime
        Public Property UpdatedAt As DateTime
        Public Property DeletedAt As Nullable(Of DateTime)

        ''' <summary>ナビゲーションプロパティ（DB には格納しない）</summary>
        Public Property Attachments As List(Of Attachment)

        Public Sub New()
            Attachments = New List(Of Attachment)()
        End Sub

    End Class

End Namespace
