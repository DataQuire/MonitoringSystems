Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Status

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Desktop\Thesis\logindata.accdb"
        Dim user As String = TextBox1.Text
        Dim password As String = TextBox2.Text
        Dim form2 As New Form2()
        Dim form6 As New Form6()

        Dim query As String = "SELECT * FROM loginform WHERE [user] = @User AND [password] = @Password"

        Try
            Using connection As New OleDbConnection(connectionString)
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@User", user)
                    command.Parameters.AddWithValue("@Password", password)

                    connection.Open()

                    Dim reader As OleDbDataReader = command.ExecuteReader()

                    If reader.Read() Then
                        form2.BringToFront()
                        form2.Show()
                        form6.TopLevel = False
                        form6.FormBorderStyle = FormBorderStyle.None
                        form6.Dock = DockStyle.Fill
                        form2.Panel4.Controls.Add(form6)
                        form6.BringToFront()
                        form6.Show()
                        Me.Hide()
                    Else
                        MsgBox("Login Failed", MsgBoxStyle.Exclamation)
                        TextBox2.Text = ""
                        TextBox1.Focus()
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Connection ERROR: " & ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
    End Sub
End Class