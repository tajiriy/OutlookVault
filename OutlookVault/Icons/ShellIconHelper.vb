Option Explicit On
Option Strict On
Option Infer Off

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace Icons

    ''' <summary>shell32.dll からアイコンを抽出してフォルダツリー用の ImageList を構築する。</summary>
    Friend NotInheritable Class ShellIconHelper

        Private Sub New()
        End Sub

        ' shell32.dll アイコンインデックス
        Private Const ShellIconFolder As Integer = 3          ' 閉じフォルダ
        Private Const ShellIconFolderOpen As Integer = 4      ' 開きフォルダ
        Private Const ShellIconTrash As Integer = 31          ' ゴミ箱（空）
        Private Const ShellIconAllFolders As Integer = 172    ' スタック／すべて

        ''' <summary>ImageList 内のインデックス定数。</summary>
        Friend Const IndexAll As Integer = 0
        Friend Const IndexFolder As Integer = 1
        Friend Const IndexFolderOpen As Integer = 2
        Friend Const IndexTrash As Integer = 3

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function ExtractIconEx(
            lpszFile As String,
            nIconIndex As Integer,
            <Out()> phiconLarge() As IntPtr,
            <Out()> phiconSmall() As IntPtr,
            nIcons As UInteger) As UInteger
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        ''' <summary>shell32.dll から小さいアイコン (16x16) を取得する。</summary>
        Private Shared Function GetShellIcon(index As Integer) As Icon
            Dim large(0) As IntPtr
            Dim small(0) As IntPtr
            ExtractIconEx("shell32.dll", index, large, small, 1UI)
            Try
                If small(0) <> IntPtr.Zero Then
                    Return CType(Icon.FromHandle(small(0)).Clone(), Icon)
                End If
            Finally
                If large(0) <> IntPtr.Zero Then DestroyIcon(large(0))
                If small(0) <> IntPtr.Zero Then DestroyIcon(small(0))
            End Try
            Return Nothing
        End Function

        ''' <summary>フォルダツリー用の ImageList を生成する。</summary>
        Friend Shared Function CreateFolderImageList() As ImageList
            Dim imgList As New ImageList()
            imgList.ColorDepth = ColorDepth.Depth32Bit
            imgList.ImageSize = New Size(16, 16)
            imgList.TransparentColor = Color.Transparent

            ' IndexAll = 0
            Dim iconAll As Icon = GetShellIcon(ShellIconAllFolders)
            If iconAll IsNot Nothing Then
                imgList.Images.Add(iconAll)
            Else
                ' フォールバック: 通常フォルダアイコン
                Dim iconFallback As Icon = GetShellIcon(ShellIconFolder)
                If iconFallback IsNot Nothing Then imgList.Images.Add(iconFallback)
            End If

            ' IndexFolder = 1
            Dim iconFolder As Icon = GetShellIcon(ShellIconFolder)
            If iconFolder IsNot Nothing Then imgList.Images.Add(iconFolder)

            ' IndexFolderOpen = 2
            Dim iconFolderOpen As Icon = GetShellIcon(ShellIconFolderOpen)
            If iconFolderOpen IsNot Nothing Then imgList.Images.Add(iconFolderOpen)

            ' IndexTrash = 3
            Dim iconTrash As Icon = GetShellIcon(ShellIconTrash)
            If iconTrash IsNot Nothing Then imgList.Images.Add(iconTrash)

            Return imgList
        End Function

    End Class

End Namespace
