Public Class FormMain
    Dim MenuStrip As New MenuStrip
    Dim TabControl As New TabControl

    Public Sub New()

        InitializeComponent()

        InitializeInterface()

    End Sub

    Private Sub InitializeInterface()
        With MenuStrip

        End With

        With TabControl
            .Dock = DockStyle.Fill
        End With

        With Me
            .Controls.Add(TabControl)
            .Controls.Add(MenuStrip)
        End With
    End Sub
End Class
