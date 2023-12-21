Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Word
Imports System.IO

Public Class Form19
    Dim dataAdapter As OleDbDataAdapter
    Dim DSet As DataSet
    Dim Tables As DataTableCollection
    Dim Source1 As BindingSource


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get the data from the input fields
        Dim data1 As String = TextBox1.Text
        Dim data2 As String = TextBox2.Text
        Dim data3 As String = ComboBox1.Text
        Dim data5 As String = TextBox4.Text
        Dim data6 As String = TextBox6.Text

        ' Check if any required field is empty
        If String.IsNullOrEmpty(data1) OrElse String.IsNullOrEmpty(data2) OrElse String.IsNullOrEmpty(data3) OrElse String.IsNullOrEmpty(data5) OrElse String.IsNullOrEmpty(data6) Then
            MessageBox.Show("Please fill in all required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Save the image file to a specific folder
        Dim imagePath As String = SaveImageFile(PictureBox1.Image)

        ' Insert the data and image file path into the database
        InsertDataToDatabase(data1, data2, data3, data5, data6, imagePath)

        MessageBox.Show("Registered Successfully!")

        RefreshDataGridView()
    End Sub
    Private Sub Form19_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ' Check if the pressed key is the forward slash ("/")
        If e.KeyChar = "/" Then
            ' Generate a random 6-digit code
            Dim random As New Random()
            Dim code As String = random.Next(100000, 999999).ToString()

            ' Insert the generated code into TextBox1
            TextBox1.Text = code
        End If
    End Sub
    Private Function SaveImageFile(image As Image) As String
        ' Create a unique file name for the image using a timestamp
        Dim fileName As String = DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".jpg"

        ' Specify the folder path to save the image
        Dim folderPath As String = "C:\Users\user\Desktop\Thesis\Images\"

        ' Combine the folder path and file name to get the full file path
        Dim filePath As String = Path.Combine(folderPath, fileName)

        ' Save the image file to the specified path
        image.Save(filePath)

        ' Return the file path
        Return filePath
    End Function
    Private Sub RefreshDataGridView()
        Dim myConnection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Desktop\Thesis\RFID.accdb;")
        DSet = New DataSet
        Tables = DSet.Tables
        dataAdapter = New OleDbDataAdapter("SELECT RFID, DriverName, Individual, Vehicle, Plate FROM RFID1", myConnection) ' Exclude ImagePath column from the SELECT statement
        dataAdapter.Fill(DSet, "RFID1")

        Source1 = New BindingSource()
        Dim View As New DataView(Tables(0))
        Source1.DataSource = View
        DataGridView1.DataSource = View
    End Sub
    Private Sub InsertDataToDatabase(data1 As String, data2 As String, data3 As String, data5 As String, data6 As String, imagePath As String)
        Dim myConnection As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Desktop\Thesis\RFID.accdb;")

        Dim insertQuery As String = "INSERT INTO RFID1 (RFID, DriverName, Individual, Vehicle, Plate, ImagePath) Values (@Data1, @Data2, @Data3, @Data5, @Data6, @ImagePath)"

        Dim insertCommand As New OleDbCommand(insertQuery, myConnection)

        insertCommand.Parameters.AddWithValue("@Data1", data1)
        insertCommand.Parameters.AddWithValue("@Data2", data2)
        insertCommand.Parameters.AddWithValue("@Data3", data3)
        insertCommand.Parameters.AddWithValue("@Data5", data5)
        insertCommand.Parameters.AddWithValue("@Data6", data6)
        insertCommand.Parameters.AddWithValue("@ImagePath", imagePath)

        myConnection.Open()

        insertCommand.ExecuteNonQuery()

        myConnection.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("RFID").Value.ToString()
            TextBox2.Text = selectedRow.Cells("DriverName").Value.ToString()
            ComboBox1.Text = selectedRow.Cells("Individual").Value.ToString()
            TextBox4.Text = selectedRow.Cells("Vehicle").Value.ToString()
            TextBox6.Text = selectedRow.Cells("Plate").Value.ToString()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Check if a row is selected in the DataGridView
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Find the corresponding row in the DataTable
            Dim drv As DataRowView = selectedRow.DataBoundItem
            Dim selectedDataRow As DataRow = drv.Row

            ' Remove the row from the DataTable
            Tables("RFID1").Rows.Remove(selectedDataRow)
        End If

    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Dim searchText As String = ComboBox2.Text.Trim()

        ' Perform the search operation based on the entered text
        Dim resultView As DataView = Tables("RFID1").DefaultView
        resultView.RowFilter = String.Format("DriverName LIKE '%{0}%'", searchText)

        ' Update the DataGridView with the search results
        DataGridView1.DataSource = resultView
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Open the OpenFileDialog to browse and select an image file
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif|All Files|*.*"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            Dim imagePath As String = openFileDialog.FileName

            ' Display the selected image in the PictureBox control without cropping
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            PictureBox1.Image = Image.FromFile(imagePath)
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form19_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RefreshDataGridView()
        TextBox1.Text = "RFID"
        TextBox1.ForeColor = Color.Gray
        TextBox2.Text = "Name"
        TextBox2.ForeColor = Color.Gray
        TextBox4.Text = "Vehicle"
        TextBox4.ForeColor = Color.Gray
        TextBox6.Text = "Plate#"
        TextBox6.ForeColor = Color.Gray
        KeyPreview = True
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "RFID" Then
            TextBox1.Text = ""
            TextBox1.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        If TextBox2.Text = "Name" Then
            TextBox2.Text = ""
            TextBox2.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        If TextBox4.Text = "Vehicle" Then
            TextBox4.Text = ""
            TextBox4.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        If TextBox6.Text = "Plate#" Then
            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Create a new Word application instance
        Dim wordApp As New Application()

        ' Create a new Word document
        Dim doc As Document = wordApp.Documents.Add()

        ' Get the data from the DataGridView
        Dim rowCount As Integer = DataGridView1.Rows.Count
        Dim columnCount As Integer = DataGridView1.Columns.Count

        ' Add a table to the document
        Dim table As Table = doc.Tables.Add(doc.Range, rowCount + 1, columnCount)

        ' Set the table headers from the DataGridView column names
        For colIndex As Integer = 0 To columnCount - 1
            table.Cell(1, colIndex + 1).Range.Text = DataGridView1.Columns(colIndex).HeaderText
        Next

        ' Set the table data from the DataGridView rows
        For rowIndex As Integer = 0 To rowCount - 1
            For colIndex As Integer = 0 To columnCount - 1
                Dim cellValue As Object = DataGridView1.Rows(rowIndex).Cells(colIndex).Value
                table.Cell(rowIndex + 2, colIndex + 1).Range.Text = If(cellValue IsNot Nothing, cellValue.ToString(), "")
            Next
        Next

        ' Save the document
        Dim filePath As String = "C:\Users\user\Desktop\Thesis\datagridview.docx"
        doc.SaveAs(filePath)

        ' Close the document and Word application
        doc.Close()
        wordApp.Quit()

        MessageBox.Show("Word document created successfully!")
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
End Class