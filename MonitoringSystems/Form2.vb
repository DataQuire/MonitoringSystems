Imports System.IO
Imports System.IO.Ports
Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
    End Sub
    Private Sub RestartForm(form As Form)
        Try
            Me.Hide()
            Dim newForm As Form = CType(Activator.CreateInstance(form.GetType()), Form)
            newForm.StartPosition = FormStartPosition.CenterScreen
            newForm.ShowDialog()
            Me.Dispose()
        Catch ex As Exception
            MessageBox.Show("An error occurred while loading the data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit the application?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.Yes Then
            Environment.Exit(0)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form20.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim form3 As New Form3()
        form3.TopLevel = False
        form3.FormBorderStyle = FormBorderStyle.None
        form3.Dock = DockStyle.Fill
        Panel4.Controls.Add(form3)
        form3.BringToFront()
        form3.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form10.BringToFront()
        Form10.Show()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim result As DialogResult = MessageBox.Show("You will be signed out?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

        If result = DialogResult.Yes Then
            Dim formsToClose As New List(Of Form)

            For Each openForm As Form In Application.OpenForms
                If openForm IsNot Me AndAlso openForm.Name <> "Form1" Then
                    formsToClose.Add(openForm)
                End If
            Next

            For Each formToClose As Form In formsToClose
                formToClose.Close()
            Next

            Form1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim form16 As New Form16()
        form16.TopLevel = False
        form16.FormBorderStyle = FormBorderStyle.None
        form16.Dock = DockStyle.Fill
        Panel4.Controls.Add(form16)
        form16.BringToFront()
        form16.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Form14.Show()
    End Sub
End Class