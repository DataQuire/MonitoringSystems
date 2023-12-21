Imports System.IO
Imports System.IO.Ports
Public Class Form5
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            SerialPort1.Close()
            SerialPort1.PortName = "COM3"
            SerialPort1.BaudRate = "9600"
            SerialPort1.DataBits = 8
            SerialPort1.Parity = Parity.None
            SerialPort1.StopBits = StopBits.One
            SerialPort1.Handshake = Handshake.None
            SerialPort1.Encoding = System.Text.Encoding.Default
            SerialPort1.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim s As String
        Dim strings As String()
        s = TextBox6.Text + "," + "," + "," + "," + "," + ""
        strings = s.Split(New Char() {","c})
        TextBox1.Text = strings(0)
        TextBox2.Text = strings(1)
        TextBox3.Text = strings(2)
        TextBox4.Text = strings(3)
        TextBox5.Text = strings(4)
        TextBox6.Text = ""
    End Sub
    Private Sub DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try
            Dim mydata As String = ""
            mydata = SerialPort1.ReadExisting()

            If TextBox6.InvokeRequired Then
                TextBox6.Invoke(DirectCast(Sub() TextBox6.Text &= mydata, MethodInvoker))
            Else
                TextBox6.Text &= mydata
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If InStr(TextBox1.Text, "P1ON") Then
            If Not CheckBox1.Checked Then
                CheckBox1.Checked = True
                TextBox1.BackColor = Color.Red
                TextBox7.Text = DateTime.Now.ToString("HH:mm:ss")
                TextBox8.Text = ""
            End If
        End If
        If InStr(TextBox1.Text, "P1OFF") Then
            If CheckBox1.Checked Then
                CheckBox1.Checked = False
                TextBox1.BackColor = Color.Green
                TextBox8.Text = DateTime.Now.ToString("HH:mm:ss")
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If InStr(TextBox2.Text, "P2ON") Then
            If Not CheckBox2.Checked Then
                CheckBox2.Checked = True
                TextBox2.BackColor = Color.Red
                TextBox9.Text = DateTime.Now.ToString("HH:mm:ss")
                TextBox10.Text = ""
            End If
        End If
        If InStr(TextBox2.Text, "P2OFF") Then
            If CheckBox2.Checked Then
                CheckBox2.Checked = False
                TextBox2.BackColor = Color.Green
                TextBox10.Text = DateTime.Now.ToString("HH:mm:ss")
            End If
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If InStr(TextBox3.Text, "P3ON") Then
            If Not CheckBox3.Checked Then
                CheckBox3.Checked = True
                TextBox3.BackColor = Color.Red
                TextBox11.Text = DateTime.Now.ToString("HH:mm:ss")
                TextBox12.Text = ""
            End If
        End If
        If InStr(TextBox3.Text, "P3OFF") Then
            If CheckBox3.Checked Then
                CheckBox3.Checked = False
                TextBox3.BackColor = Color.Green
                TextBox12.Text = DateTime.Now.ToString("HH:mm:ss")
            End If
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If InStr(TextBox4.Text, "P4ON") Then
            If Not CheckBox4.Checked Then
                CheckBox4.Checked = True
                TextBox4.BackColor = Color.Red
                TextBox13.Text = DateTime.Now.ToString("HH:mm:ss")
                TextBox14.Text = ""
            End If
        End If
        If InStr(TextBox4.Text, "P4OFF") Then
            If CheckBox4.Checked Then
                CheckBox4.Checked = False
                TextBox4.BackColor = Color.Green
                TextBox14.Text = DateTime.Now.ToString("HH:mm:ss")
            End If
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If InStr(TextBox5.Text, "P5ON") Then
            If Not CheckBox5.Checked Then
                CheckBox5.Checked = True
                TextBox5.BackColor = Color.Red
                TextBox15.Text = DateTime.Now.ToString("HH:mm:ss")
                TextBox16.Text = ""
            End If
        End If
        If InStr(TextBox5.Text, "P5OFF") Then
            If CheckBox5.Checked Then
                CheckBox5.Checked = False
                TextBox5.BackColor = Color.Green
                TextBox16.Text = DateTime.Now.ToString("HH:mm:ss")
            End If
        End If
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text <> "" Then
            Dim textbox7Content As String = TextBox7.Text
            Dim textbox8Content As String = TextBox8.Text
            Dim message As String = $"PARKING SLOT 1;  Time Arrival/Departure: {textbox7Content}, {textbox8Content}"
            ListBox1.Items.Add(message)
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text <> "" Then
            Dim textbox9Content As String = TextBox9.Text
            Dim textbox10Content As String = TextBox10.Text
            Dim message As String = $"PARKING SLOT 2;  Time Arrival/Departure: {textbox9Content}, {textbox10Content}"
            ListBox1.Items.Add(message)
        End If
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        If TextBox12.Text <> "" Then
            Dim textbox11Content As String = TextBox11.Text
            Dim textbox12Content As String = TextBox12.Text
            Dim message As String = $"PARKING SLOT 3;  Time Arrival/Departure: {textbox11Content}, {textbox12Content}"
            ListBox1.Items.Add(message)
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If TextBox14.Text <> "" Then
            Dim textbox13Content As String = TextBox13.Text
            Dim textbox14Content As String = TextBox14.Text
            Dim message As String = $"PARKING SLOT 4;  Time Arrival/Departure: {textbox13Content}, {textbox14Content}"
            ListBox1.Items.Add(message)
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        If TextBox16.Text <> "" Then
            Dim textbox15Content As String = TextBox15.Text
            Dim textbox16Content As String = TextBox16.Text
            Dim message As String = $"PARKING SLOT 5;  Time Arrival/Departure: {textbox15Content}, {textbox16Content}"
            ListBox1.Items.Add(message)
        End If
    End Sub
    Private filePath As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            filePath = saveFileDialog.FileName
            SaveDataToFile(ListBox1.Items)
            MessageBox.Show($"Data saved to: {filePath}", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Private Sub SaveDataToFile(items As ListBox.ObjectCollection)
        Using writer As New StreamWriter(filePath)
            For Each item As Object In items
                writer.WriteLine(item.ToString())
            Next
        End Using
    End Sub
End Class