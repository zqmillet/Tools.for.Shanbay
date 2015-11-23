Public Class FormMain
    Dim MenuStrip As New This.MenuStrip
    Dim TabControl As New This.TabControl

    Public Sub New()
        InitializeComponent()
        InitializeInterface()
    End Sub

    Private Sub InitializeInterface()
        With Me
            .Controls.Add(TabControl)
            .Controls.Add(MenuStrip)
        End With
    End Sub
End Class
