Public Class Form16
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel2.Controls.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form15 As New Form15()
        form15.TopLevel = False
        form15.FormBorderStyle = FormBorderStyle.None
        form15.Dock = DockStyle.Fill
        Panel2.Controls.Add(form15)
        form15.BringToFront()
        form15.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim form19 As New Form19()
        form19.TopLevel = False
        form19.FormBorderStyle = FormBorderStyle.None
        form19.Dock = DockStyle.Fill
        Panel2.Controls.Add(form19)
        form19.BringToFront()
        form19.Show()
    End Sub
End Class