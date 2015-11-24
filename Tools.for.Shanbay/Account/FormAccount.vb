Imports System.Text.RegularExpressions
Imports HtmlAgilityPack

Public Class FormAccount
    Private AccountFilePath As String = Application.StartupPath & "\Account.nfo"
    Private RowHeight As Integer = 25
    Private FormHeightWithCaptcha As Integer = 138
    Private FormHeightWithoutCaptcha As Integer = FormHeightWithCaptcha - RowHeight - 2

    Private TableLayoutPanel As New TableLayoutPanel
    Private ButtonTest As New Button
    Private ButtonSave As New Button
    Private ButtonQuit As New Button
    Private TextBoxUserName As New TextBox
    Private TextBoxPassword As New TextBox
    Private TextBoxCaptcha As New TextBox
    Private LabelUserName As New Label
    Private LabelPassword As New Label
    Private LabelCaptcha As New Label
    Private PictureBoxCaptcha As New PictureBox

    Private WebBrowser As New WebBrowser
    Private StatusStrip As New StatusStrip
    Private ToolStripStatusLabel As New ToolStripStatusLabel

    Private WebBrowserLoaded As Boolean = False

    Public Sub New()
        InitializeComponent()
        InitializeInterface()
        InitializeWebBrowser()
        ReadAccountInformation()
    End Sub

    Private Sub InitializeWebBrowser()
        With WebBrowser
            AddHandler WebBrowser.DocumentCompleted, AddressOf WebBrowser_DocumentCompleted
            .Navigate("http://www.shanbay.com/accounts/login/")
            .Dock = DockStyle.Fill
            Do Until .ReadyState = WebBrowserReadyState.Complete
                Application.DoEvents()
            Loop
        End With

        ShowCaptchaLine()
    End Sub

    Private Sub WebBrowser_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        WebBrowserLoaded = True
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

        With LabelCaptcha
            .Text = "Captcha"
            .Visible = False
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

        With TextBoxCaptcha
            .Dock = DockStyle.Fill
            .Visible = False
            AddHandler .TextChanged, AddressOf TextBox_TextChanged
        End With

        With PictureBoxCaptcha
            .Dock = DockStyle.Fill
            .Visible = False
            .SizeMode = PictureBoxSizeMode.StretchImage
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
            .ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 60.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))
            .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0F))

            .RowCount = 4
            .RowStyles.Add(New RowStyle(SizeType.Absolute, RowHeight))
            .RowStyles.Add(New RowStyle(SizeType.Absolute, RowHeight))
            .RowStyles.Add(New RowStyle(SizeType.Absolute, RowHeight))
            .RowStyles.Add(New RowStyle(SizeType.Absolute, RowHeight))
            .Padding = New Padding(2)

            .Controls.Add(LabelUserName, 0, 0)
            .SetColumnSpan(TextBoxUserName, 3)
            .Controls.Add(TextBoxUserName, 1, 0)

            .Controls.Add(LabelPassword, 0, 1)
            .SetColumnSpan(TextBoxPassword, 3)
            .Controls.Add(TextBoxPassword, 1, 1)

            .Controls.Add(LabelCaptcha, 0, 2)
            .SetColumnSpan(TextBoxCaptcha, 2)
            .Controls.Add(TextBoxCaptcha, 1, 2)
            .Controls.Add(PictureBoxCaptcha, 3, 2)
            .RowStyles(2).Height = 0

            .Controls.Add(ButtonTest, 1, 3)
            .Controls.Add(ButtonSave, 2, 3)
            .Controls.Add(ButtonQuit, 3, 3)
        End With

        With ToolStripStatusLabel
            .Text = ""
        End With

        With StatusStrip
            .SizingGrip = False
            .Items.Add(ToolStripStatusLabel)
        End With

        With Me
            .ClientSize = New Size(300, FormHeightWithoutCaptcha)
            .FormBorderStyle = FormBorderStyle.FixedDialog

            .ControlBox = False
            .MaximizeBox = False

            .Controls.Add(TableLayoutPanel)
            .Controls.Add(StatusStrip)
            .AcceptButton = ButtonSave
            .CancelButton = ButtonQuit
        End With
    End Sub

    Private Sub ShowCaptchaLine()
        Dim HtmlDocument As New HtmlAgilityPack.HtmlDocument
        HtmlDocument.LoadHtml(WebBrowser.Document.Body.InnerHtml)

        Dim Visible = Not (HtmlDocument.DocumentNode.SelectSingleNode("//input[@id='id_captcha_1']") Is Nothing)
        If Visible Then
            Me.ClientSize = New Size(Me.ClientSize.Width, FormHeightWithCaptcha)
            TableLayoutPanel.RowStyles(2).Height = RowHeight
        Else
            Me.ClientSize = New Size(Me.ClientSize.Width, FormHeightWithoutCaptcha)
            TableLayoutPanel.RowStyles(2).Height = 0
        End If

        LabelCaptcha.Visible = Visible
        TextBoxCaptcha.Visible = Visible
        PictureBoxCaptcha.Visible = Visible

        If Visible Then
            Dim CaptchaImageURL As String = ""
            For Each HtmlNode As HtmlNode In HtmlDocument.DocumentNode.SelectNodes("//img[@class='captcha']")
                CaptchaImageURL = HtmlNode.GetAttributeValue("src", "")
            Next
            If CaptchaImageURL = "" Then
                MsgBox("Can't Find Captcha Image URL")
            Else
                PictureBoxCaptcha.ImageLocation = CaptchaImageURL
            End If
        End If
    End Sub

    Private Sub ButtonTest_Click(sender As Object, e As EventArgs)
        ButtonTest.Enabled = False

        Dim TextBoxUsername As HtmlElement = WebBrowser.Document.GetElementById("id_username")
        If TextBoxUsername Is Nothing Then
            Exit Sub
        End If

        Dim TextBoxPassword As HtmlElement = WebBrowser.Document.GetElementById("id_password")
        If TextBoxPassword Is Nothing Then
            Exit Sub
        End If

        If Me.TextBoxCaptcha.Visible Then
            Dim TextBoxCaptcha As HtmlElement = WebBrowser.Document.GetElementById("id_captcha_1")
            If TextBoxCaptcha Is Nothing Then
                Exit Sub
            End If
            TextBoxCaptcha.SetAttribute("value", Me.TextBoxCaptcha.Text)
        End If

        Dim ButtonSubmit As HtmlElement = Nothing
        For Each HtmlElement As HtmlElement In WebBrowser.Document.GetElementsByTagName("button")
            If HtmlElement.OuterHtml.Replace(" ", "").Contains("type=submit") Then
                ButtonSubmit = HtmlElement
            End If
        Next
        If ButtonSubmit Is Nothing Then
            Exit Sub
        End If

        TextBoxUsername.SetAttribute("value", Me.TextBoxUserName.Text)
        TextBoxPassword.SetAttribute("value", Me.TextBoxPassword.Text)

        WebBrowserLoaded = False
        ButtonSubmit.InvokeMember("click")
        While Not WebBrowserLoaded
            Application.DoEvents()
        End While

        If WebBrowser.Document.GetElementById("id_username") Is Nothing Then
            ToolStripStatusLabel.Text = "The account and password are correct!"
        Else
            ToolStripStatusLabel.Text = "The account and password are not correct!"
        End If

        ButtonTest.Enabled = True
        InitializeWebBrowser()
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