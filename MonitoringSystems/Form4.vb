
Public Class Form4
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

End Class