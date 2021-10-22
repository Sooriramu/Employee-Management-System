Public Class Form9
    Private Sub Form9_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        Form1.Show()

    End Sub

    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
    End Sub
End Class