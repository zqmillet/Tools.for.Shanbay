Namespace This
    Public Class MenuStrip
        Inherits Windows.Forms.MenuStrip

        Dim MenuItem_Start As New ToolStripMenuItem
        Dim MenuItem_Account As New ToolStripMenuItem

        Public Event MenuItem_Account_Click()


        Public Sub New()
            With MenuItem_Account
                .Text = "Account Configuration"
                AddHandler .Click, AddressOf ToolStripMenuItem_Click
            End With

            With MenuItem_Start
                .Text = "Start"
                .Name = "MenuItem_Start"
                .DropDownItems.Add(MenuItem_Account)
            End With

            With Me
                .Items.Add(MenuItem_Start)
            End With
        End Sub

        Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Select Case CType(sender, ToolStripMenuItem)
                Case MenuItem_Account
                    RaiseEvent MenuItem_Account_Click()
            End Select
        End Sub
    End Class

    Public Class ToolStripMenuItem
        Inherits Windows.Forms.ToolStripMenuItem

        Public Shared Operator =(ByVal ToolStripMenuItem1 As ToolStripMenuItem,
                                 ByVal ToolStripMenuItem2 As ToolStripMenuItem) As Boolean
            Return ToolStripMenuItem1 Is ToolStripMenuItem2
        End Operator

        Public Shared Operator <>(ByVal ToolStripMenuItem1 As ToolStripMenuItem,
                                 ByVal ToolStripMenuItem2 As ToolStripMenuItem) As Boolean
            Return Not (ToolStripMenuItem1 Is ToolStripMenuItem2)
        End Operator
    End Class
End Namespace

