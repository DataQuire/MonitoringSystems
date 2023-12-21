Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Status

Public Class Form13
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
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
End Class