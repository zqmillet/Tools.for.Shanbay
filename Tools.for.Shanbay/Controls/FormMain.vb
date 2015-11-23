Public Class FormMain
    Dim MenuStrip As New _FormMain.MenuStrip
    Dim TabControl As New _FormMain.TabControl

    Public Sub New()
        InitializeComponent()
        InitializeInterface()
    End Sub

    Private Sub InitializeInterface()
        With MenuStrip
            AddHandler .MenuItem_Account_Click, AddressOf MenuStrip_MenuItem_Account_Click

        End With

        With Me
            .Controls.Add(TabControl)
            .Controls.Add(MenuStrip)
        End With

        MenuStrip_MenuItem_Account_Click()
    End Sub

    Private Sub MenuStrip_MenuItem_Account_Click()
        Dim FormAccount As New FormAccount
        FormAccount.ShowDialog()
    End Sub

End Class
