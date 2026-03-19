Option Explicit On
Option Strict On
Option Infer Off

Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms

Namespace Forms

    Public Class AboutForm
        Inherits Form

        Private picIcon As PictureBox
        Private lblAppName As Label
        Private lblVersion As Label
        Private lblDescription As Label
        Private btnOk As Button

        Public Sub New()
            Me.Text = "バージョン情報"
            Me.Size = New Size(340, 200)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False

            picIcon = New PictureBox()
            picIcon.Location = New Point(24, 24)
            picIcon.Size = New Size(64, 64)
            picIcon.SizeMode = PictureBoxSizeMode.Zoom

            Dim icoPath As String = System.IO.Path.Combine(Application.StartupPath, "app.ico")
            If System.IO.File.Exists(icoPath) Then
                Dim ico As New Icon(icoPath, 256, 256)
                picIcon.Image = ico.ToBitmap()
            End If

            lblAppName = New Label()
            lblAppName.Text = "OutlookVault"
            lblAppName.Font = New Font(Me.Font.FontFamily, 14.0F, FontStyle.Bold)
            lblAppName.Location = New Point(104, 24)
            lblAppName.AutoSize = True

            Dim ver As Version = Assembly.GetExecutingAssembly().GetName().Version
            lblVersion = New Label()
            lblVersion.Text = "Version " & ver.Major.ToString() & "." & ver.Minor.ToString() & "." & ver.Build.ToString()
            lblVersion.Location = New Point(104, 56)
            lblVersion.AutoSize = True

            lblDescription = New Label()
            lblDescription.Text = "Outlook メール保管ツール"
            lblDescription.Location = New Point(104, 80)
            lblDescription.AutoSize = True

            btnOk = New Button()
            btnOk.Text = "OK"
            btnOk.Size = New Size(80, 28)
            btnOk.Location = New Point(124, 130)
            AddHandler btnOk.Click, Sub(s As Object, ev As EventArgs)
                                        Me.Close()
                                    End Sub
            Me.AcceptButton = btnOk

            Me.Controls.Add(picIcon)
            Me.Controls.Add(lblAppName)
            Me.Controls.Add(lblVersion)
            Me.Controls.Add(lblDescription)
            Me.Controls.Add(btnOk)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If picIcon IsNot Nothing AndAlso picIcon.Image IsNot Nothing Then
                    picIcon.Image.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

    End Class

End Namespace
