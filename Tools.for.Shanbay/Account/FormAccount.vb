Imports System.Text.RegularExpressions

Public Class FormAccount
    Private AccountFilePath As String = Application.StartupPath & "\Account.nfo"

    Private TableLayoutPanel As New TableLayoutPanel
    Private ButtonTest As New Button
    Private ButtonSave As New Button
    Private ButtonQuit As New Button
    Private TextBoxUserName As New TextBox
    Private TextBoxPassword As New TextBox
    Private LabelUserName As New Label
    Private LabelPassword As New Label

    Private WebBrowser As New WebBrowser

    Public Sub New()
        InitializeComponent()
        InitializeInterface()
        ReadAccountInformation()
    End Sub

    Private Sub InitializeInterface()
        With LabelUserName
            .Text = "UserName"
            .TextAlign = ContentAlignment.MiddleLeft
        End With

        With LabelPassword
            .Text = "Password"
            .TextAlign = ContentAlignment.MiddleLeft
        End With

        With TextBoxUserName
            .Dock = DockStyle.Fill
            AddHandler .TextChanged, AddressOf TextBox_TextChanged
        End With

        With TextBoxPassword
            .Dock = DockStyle.Fill
            .PasswordChar = "*"
            AddHandler .TextChanged, AddressOf TextBox_TextChanged
        End With

        With ButtonTest
            .Text = "Test"
            .Dock = DockStyle.Fill
            .Enabled = False
            AddHandler .Click, AddressOf ButtonTest_Click
        End With

        With ButtonSave
            .Text = "Save"
            .Dock = DockStyle.Fill
            .Enabled = False
            AddHandler .Click, AddressOf ButtonSave_Click
        End With

        With ButtonQuit
            .Text = "Cancel"
            .Dock = DockStyle.Fill
            AddHandler .Click, AddressOf ButtonQuit_Click
        End With

        With TableLayoutPanel
            .Dock = DockStyle.Fill
            .ColumnCount = 4
            .RowCount = 3
            .ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 60.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            .Padding = New Padding(2)

            .Controls.Add(LabelUserName, 0, 0)
            .SetColumnSpan(TextBoxUserName, 3)
            .Controls.Add(TextBoxUserName, 1, 0)

            .Controls.Add(LabelPassword, 0, 1)
            .SetColumnSpan(TextBoxPassword, 3)
            .Controls.Add(TextBoxPassword, 1, 1)

            .Controls.Add(ButtonTest, 1, 2)
            .Controls.Add(ButtonSave, 2, 2)
            .Controls.Add(ButtonQuit, 3, 2)
        End With

        With Me
            .Size = New Size(300, 130)
            .FormBorderStyle = FormBorderStyle.FixedSingle

            .ControlBox = False
            .Controls.Add(TableLayoutPanel)

            .AcceptButton = ButtonSave
            .CancelButton = ButtonQuit
        End With
    End Sub

    Private Sub ButtonTest_Click(sender As Object, e As EventArgs)
        WebBrowser.Navigate("http://www.shanbay.com/accounts/login/")

    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs)
        Dim Writer As New IO.StreamWriter(AccountFilePath, False, System.Text.Encoding.Default)

        Writer.WriteLine(TextBoxUserName.Text)
        Writer.WriteLine(TextBoxPassword.Text)

        Writer.Close()
        Me.Dispose()
    End Sub

    Private Sub ButtonQuit_Click(sender As Object, e As EventArgs)
        Me.Dispose()
    End Sub

    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs)
        Dim Enable As Boolean = Regex.IsMatch(TextBoxUserName.Text, "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*") And
            TextBoxPassword.Text <> ""

        ButtonTest.Enabled = Enable
        ButtonSave.Enabled = Enable
    End Sub

    Private Sub ReadAccountInformation()
        If Not My.Computer.FileSystem.FileExists(AccountFilePath) Then
            MsgBox("Can't File Account.acc!", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        Dim Reader As New IO.StreamReader(AccountFilePath, System.Text.Encoding.Default)

        TextBoxUserName.Text = Reader.ReadLine
        TextBoxPassword.Text = Reader.ReadLine

        Reader.Close()
    End Sub
End Class