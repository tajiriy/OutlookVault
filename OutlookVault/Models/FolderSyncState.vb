Option Explicit On
Option Strict On
Option Infer Off

Namespace Models

    ''' <summary>
    ''' フォルダごとの同期状態を保持するモデルクラス。
    ''' 差分スキャンの基準日時と、完全同期の完了状態を管理する。
    ''' </summary>
    Public Class FolderSyncState

        ''' <summary>フォルダの表示名</summary>
        Public Property FolderName As String

        ''' <summary>最後に同期を完了した日時</summary>
        Public Property LastSyncTime As DateTime

        ''' <summary>完全同期が1回以上完了しているか</summary>
        Public Property FullSyncDone As Boolean

    End Class

End Namespace
