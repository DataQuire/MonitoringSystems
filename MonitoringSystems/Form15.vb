Imports System.Data.OleDb
Public Class Form15
    Dim conn As New OleDbConnection
    Dim cmd As OleDbCommand
    Dim dt As New DataTable
    Dim da As New OleDbDataAdapter(cmd)

    Private bitmap As Bitmap
    Private Sub Form15_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Desktop\Thesis\rooms.accdb"
        view()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "insert into roomm(Building, Room, Floor) values ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "')"
            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Save successful", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update roomm set Room '" + TextBox5.Text + "' where Building = '" + TextBox4.Text + "' and Floor = '" + TextBox6.Text + "'"
            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Save successful", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Dim selectedRow = DataGridView1.Rows(e.RowIndex)
            TextBox1.Text = selectedRow.Cells(0).Value.ToString()
            TextBox2.Text = selectedRow.Cells(1).Value.ToString()
            TextBox3.Text = selectedRow.Cells(2).Value.ToString()
            TextBox4.Text = selectedRow.Cells(0).Value.ToString()
            TextBox5.Text = selectedRow.Cells(1).Value.ToString()
            TextBox6.Text = selectedRow.Cells(2).Value.ToString()
            TextBox8.Text = selectedRow.Cells(0).Value.ToString()
            TextBox9.Text = selectedRow.Cells(1).Value.ToString()
            TextBox10.Text = selectedRow.Cells(2).Value.ToString()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub view()
        conn.Open()
        cmd = conn.CreateCommand()
        cmd.CommandType = CommandType.Text
        da = New OleDbDataAdapter("select * From roomm", conn)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        conn.Close()

        DataGridView1.Columns(0).Width = 300
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).Width = 200
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim searchTerm As String = TextBox7.Text.Trim().ToLower()
        Dim matchFound As Boolean = False

        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim rowContainsTerm As Boolean = False

            For Each cell As DataGridViewCell In row.Cells
                If cell.Value IsNot Nothing AndAlso cell.Value.ToString().ToLower().Contains(searchTerm) Then
                    rowContainsTerm = True
                    cell.Selected = True
                    Exit For
                End If
            Next
            If rowContainsTerm Then
                row.Selected = True
                DataGridView1.FirstDisplayedScrollingRowIndex = row.Index
                matchFound = True
            End If
        Next
        If Not matchFound Then
            MessageBox.Show("No matches found.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
            conn.Open()
            cmd = conn.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE * FROM roomm WHERE Room = '" + TextBox9.Text + "' AND Building = '" + TextBox8.Text + "' AND Floor = '" + TextBox10.Text + "'"

            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Successfully Deleted!", MsgBoxStyle.Information)
            Button4_Click(New Object, New EventArgs())
            view()
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Vb Save Database", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()

        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        view()
    End Sub
End Class


