Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Status
Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
    End Sub
    Private Sub RestartForm(form As Form)
        Try
            ' Clean up any resources or connections associated with the form
            ' ...

            ' Hide the current form
            Me.Hide()

            ' Recreate a new instance of the form
            Dim newForm As Form = CType(Activator.CreateInstance(form.GetType()), Form)
            newForm.StartPosition = FormStartPosition.CenterScreen

            ' Show the new form
            newForm.ShowDialog()

            ' Dispose the current form after the new form is closed
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show("An error occurred while loading the data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim form4 As New Form4()
        form4.TopLevel = False
        form4.FormBorderStyle = FormBorderStyle.None
        form4.Dock = DockStyle.Fill
        Panel2.Controls.Add(form4)
        form4.BringToFront()
        form4.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim form7 As New Form7()
        form7.TopLevel = False
        form7.FormBorderStyle = FormBorderStyle.None
        form7.Dock = DockStyle.Fill
        Panel2.Controls.Add(form7)
        form7.BringToFront()
        form7.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim form8 As New Form8()
        form8.TopLevel = False
        form8.FormBorderStyle = FormBorderStyle.None
        form8.Dock = DockStyle.Fill
        Panel2.Controls.Add(form8)
        form8.BringToFront()
        form8.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim form9 As New Form9()
        form9.TopLevel = False
        form9.FormBorderStyle = FormBorderStyle.None
        form9.Dock = DockStyle.Fill
        Panel2.Controls.Add(form9)
        form9.BringToFront()
        form9.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim form12 As New Form12()
        form12.TopLevel = False
        form12.FormBorderStyle = FormBorderStyle.None
        form12.Dock = DockStyle.Fill
        Panel2.Controls.Add(form12)
        form12.BringToFront()
        form12.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form11 As New Form11()
        form11.TopLevel = False
        form11.FormBorderStyle = FormBorderStyle.None
        form11.Dock = DockStyle.Fill
        Panel2.Controls.Add(form11)
        form11.BringToFront()
        form11.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel2.Controls.Clear()
    End Sub
End Class