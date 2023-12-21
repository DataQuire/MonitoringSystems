Imports System.Data.OleDb

Public Class Form20
    Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Desktop\Thesis\RFID.accdb"
    Private timeIn As DateTime ' Variable to store the time-in value
    Private oddEvenCounter As Integer = 0 ' Counter to keep track of odd and even times

    Private Sub Form20_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = "/" Then
            ' Connect to the database
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Check if the record already exists in RFID2 table

                ' Prepare the SQL query to retrieve driver information
                Dim query As String = "SELECT RFID, DriverName, Individual, Vehicle, Plate, ImagePath FROM RFID1 WHERE DriverName = @DriverName"
                    Using command As New OleDbCommand(query, connection)
                        ' Set the parameter value
                        command.Parameters.AddWithValue("@DriverName", "Mark Joseph Corpin")

                    ' Execute the query and retrieve the data
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Display the data in the text boxes
                        TextBox1.Text = reader("RFID").ToString()
                        TextBox2.Text = reader("DriverName").ToString()
                        TextBox3.Text = reader("Individual").ToString()
                        TextBox4.Text = reader("Vehicle").ToString()
                        TextBox5.Text = reader("Plate").ToString()

                        ' Retrieve and display the image
                        Dim imagePath As String = reader("ImagePath").ToString()
                        If Not String.IsNullOrWhiteSpace(imagePath) Then
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                            PictureBox1.Image = Image.FromFile(imagePath)
                        Else
                            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                            PictureBox1.Image = Nothing
                        End If

                        ' Increment the counter
                        oddEvenCounter += 1

                        ' Check if it's an odd or even time
                        If oddEvenCounter Mod 2 = 1 Then
                            ' Odd time - Insert time-in value
                            timeIn = DateTime.Now
                            InsertDataIntoRFID2(connection, reader("RFID").ToString(), reader("DriverName").ToString(), reader("Individual").ToString(), timeIn.ToString("yyyy-MM-dd HH:mm:ss"))
                        Else
                            ' Even time - Insert time-out value
                            Dim timeOut As DateTime = DateTime.Now
                            Dim timeDifference As TimeSpan = timeOut - timeIn
                            InsertTimeOutInRFID2(connection, reader("DriverName").ToString(), timeOut.ToString("yyyy-MM-dd HH:mm:ss"), timeDifference.TotalMinutes)
                        End If

                        UpdateDataGridView(connection)
                    Else
                        ' Name not found
                        TextBox1.Text = "Name not found"
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                        PictureBox1.Image = Nothing
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub Form20_KeyPress1(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = "." Then
            ' Connect to the database
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Check if the record already exists in RFID2 table

                ' Prepare the SQL query to retrieve driver information
                Dim query As String = "SELECT RFID, DriverName, Individual, Vehicle, Plate, ImagePath FROM RFID1 WHERE DriverName = @DriverName"
                    Using command As New OleDbCommand(query, connection)
                        ' Set the parameter value
                        command.Parameters.AddWithValue("@DriverName", "Vince Salloman")

                    ' Execute the query and retrieve the data
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Display the data in the text boxes
                        TextBox1.Text = reader("RFID").ToString()
                        TextBox2.Text = reader("DriverName").ToString()
                        TextBox3.Text = reader("Individual").ToString()
                        TextBox4.Text = reader("Vehicle").ToString()
                        TextBox5.Text = reader("Plate").ToString()

                        ' Retrieve and display the image
                        Dim imagePath As String = reader("ImagePath").ToString()
                        If Not String.IsNullOrWhiteSpace(imagePath) Then
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                            PictureBox1.Image = Image.FromFile(imagePath)
                        Else
                            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                            PictureBox1.Image = Nothing
                        End If

                        ' Increment the counter
                        oddEvenCounter += 1

                        ' Check if it's an odd or even time
                        If oddEvenCounter Mod 2 = 1 Then
                            ' Odd time - Insert time-in value
                            timeIn = DateTime.Now
                            InsertDataIntoRFID2(connection, reader("RFID").ToString(), reader("DriverName").ToString(), reader("Individual").ToString(), timeIn.ToString("yyyy-MM-dd HH:mm:ss"))
                        Else
                            ' Even time - Insert time-out value
                            Dim timeOut As DateTime = DateTime.Now
                            Dim timeDifference As TimeSpan = timeOut - timeIn
                            InsertTimeOutInRFID2(connection, reader("DriverName").ToString(), timeOut.ToString("yyyy-MM-dd HH:mm:ss"), timeDifference.TotalMinutes)
                        End If

                        UpdateDataGridView(connection)
                    Else
                        ' Name not found
                        TextBox1.Text = "Name not found"
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                        PictureBox1.Image = Nothing
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub Form20_KeyPress2(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = "," Then
            ' Connect to the database
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Prepare the SQL query to retrieve driver information
                Dim query As String = "SELECT RFID, DriverName, Individual, Vehicle, Plate, ImagePath FROM RFID1 WHERE DriverName = @DriverName"
                Using command As New OleDbCommand(query, connection)
                    ' Set the parameter value
                    command.Parameters.AddWithValue("@DriverName", "Gilmer J Castin")

                    ' Execute the query and retrieve the data
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Display the data in the text boxes
                        TextBox1.Text = reader("RFID").ToString()
                        TextBox2.Text = reader("DriverName").ToString()
                        TextBox3.Text = reader("Individual").ToString()
                        TextBox4.Text = reader("Vehicle").ToString()
                        TextBox5.Text = reader("Plate").ToString()

                        ' Retrieve and display the image
                        Dim imagePath As String = reader("ImagePath").ToString()
                        If Not String.IsNullOrWhiteSpace(imagePath) Then
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                            PictureBox1.Image = Image.FromFile(imagePath)
                        Else
                            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                            PictureBox1.Image = Nothing
                        End If

                        ' Increment the counter
                        oddEvenCounter += 1

                        ' Check if it's an odd or even time
                        If oddEvenCounter Mod 2 = 1 Then
                            ' Odd time - Insert time-in value
                            timeIn = DateTime.Now
                            InsertDataIntoRFID2(connection, reader("RFID").ToString(), reader("DriverName").ToString(), reader("Individual").ToString(), timeIn.ToString("yyyy-MM-dd HH:mm:ss"))
                        Else
                            ' Even time - Insert time-out value
                            Dim timeOut As DateTime = DateTime.Now
                            Dim timeDifference As TimeSpan = timeOut - timeIn
                            InsertTimeOutInRFID2(connection, reader("DriverName").ToString(), timeOut.ToString("yyyy-MM-dd HH:mm:ss"), timeDifference.TotalMinutes)
                        End If

                        UpdateDataGridView(connection)
                    Else
                        ' Name not found
                        TextBox1.Text = "Name not found"
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                        PictureBox1.Image = Nothing
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub InsertDataIntoRFID2(connection As OleDbConnection, rfid As String, driverName As String, individual As String, timeStamp As String)
        ' Prepare the SQL query to insert data into RFID2 table
        Dim query As String = "INSERT INTO RFID2 (RFID, DriverName, Individual, TimeIn) VALUES (@RFID, @DriverName, @Individual, @TimeStamp)"
        Using command As New OleDbCommand(query, connection)
            ' Set the parameter values
            command.Parameters.AddWithValue("@RFID", rfid)
            command.Parameters.AddWithValue("@DriverName", driverName)
            command.Parameters.AddWithValue("@Individual", individual)
            command.Parameters.AddWithValue("@TimeStamp", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

            ' Execute the query to insert the data
            command.ExecuteNonQuery()
        End Using
    End Sub
    Private Sub InsertTimeOutInRFID2(connection As OleDbConnection, driverName As String, timeOut As String, timeDifference As Double)
        ' Prepare the SQL query to update time-out and time difference in RFID2 table
        Dim query As String = "UPDATE RFID2 SET TimeOut = @TimeOut, TimeDifference = @TimeDifference WHERE DriverName = @DriverName AND TimeOut IS NULL"
        Using command As New OleDbCommand(query, connection)
            ' Set the parameter values
            command.Parameters.AddWithValue("@TimeOut", timeOut)
            command.Parameters.AddWithValue("@TimeDifference", timeDifference)
            command.Parameters.AddWithValue("@DriverName", driverName)

            ' Execute the query to update the data
            command.ExecuteNonQuery()
        End Using
    End Sub
    Private Sub UpdateDataGridView(connection As OleDbConnection)
        ' Prepare the SQL query to retrieve all records from the RFID2 table
        Dim query As String = "SELECT RFID, DriverName, Individual, TimeIn, TimeOut FROM RFID2"
        Using command As New OleDbCommand(query, connection)
            ' Execute the query and retrieve the data
            Dim adapter As New OleDbDataAdapter(command)
            Dim dataTable As New DataTable()
            adapter.Fill(dataTable)

            ' Update the DataGridView
            DataGridView1.DataSource = dataTable
            DataGridView1.Columns("TimeIn").DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss"
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End Using
    End Sub
    Private Sub Form20_KeyPress3(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If e.KeyChar = "'" Then
            ' Connect to the database
            Using connection As New OleDbConnection(connectionString)
                connection.Open()

                ' Prepare the SQL query to retrieve driver information
                ' Check if the record already exists in RFID2 table
                ' Prepare the SQL query to retrieve driver information
                Dim query As String = "SELECT RFID, DriverName, Individual, Vehicle, Plate, ImagePath FROM RFID1 WHERE DriverName = @DriverName"
                Using command As New OleDbCommand(query, connection)
                    ' Set the parameter value
                    command.Parameters.AddWithValue("@DriverName", "Christopher Gregory")

                    ' Execute the query and retrieve the data
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Display the data in the text boxes
                        TextBox1.Text = reader("RFID").ToString()
                        TextBox2.Text = reader("DriverName").ToString()
                        TextBox3.Text = reader("Individual").ToString()
                        TextBox4.Text = reader("Vehicle").ToString()
                        TextBox5.Text = reader("Plate").ToString()

                        ' Retrieve and display the image
                        Dim imagePath As String = reader("ImagePath").ToString()
                        If Not String.IsNullOrWhiteSpace(imagePath) Then
                            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                            PictureBox1.Image = Image.FromFile(imagePath)
                        Else
                            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                            PictureBox1.Image = Nothing
                        End If

                        ' Increment the counter
                        oddEvenCounter += 1

                        ' Check if it's an odd or even time
                        If oddEvenCounter Mod 2 = 1 Then
                            ' Odd time - Insert time-in value
                            timeIn = DateTime.Now
                            InsertDataIntoRFID2(connection, reader("RFID").ToString(), reader("DriverName").ToString(), reader("Individual").ToString(), timeIn.ToString("yyyy-MM-dd HH:mm:ss"))
                        Else
                            ' Even time - Insert time-out value
                            Dim timeOut As DateTime = DateTime.Now
                            Dim timeDifference As TimeSpan = timeOut - timeIn
                            InsertTimeOutInRFID2(connection, reader("DriverName").ToString(), timeOut.ToString("yyyy-MM-dd HH:mm:ss"), timeDifference.TotalMinutes)
                        End If

                        UpdateDataGridView(connection)
                    Else
                        ' Name not found
                        TextBox1.Text = "Name not found"
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
                        PictureBox1.Image = Nothing
                    End If
                End Using
            End Using
        End If
    End Sub

    Private Sub Form20_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        KeyPreview = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub
End Class