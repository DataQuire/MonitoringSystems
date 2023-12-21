Imports System.Data.OleDb
Imports System.Net

Public Class Form10
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        viewer()
    End Sub
    Private Sub Checkbox1_CheckedChanged(sender As Object, e As EventArgs) Handles Checkbox1.CheckedChanged
        If Checkbox1.Checked Then
            Checkbox2.Checked = False
        End If
    End Sub
    Private Sub Checkbox2_CheckedChanged(sender As Object, e As EventArgs) Handles Checkbox2.CheckedChanged
        If Checkbox2.Checked Then
            Checkbox1.Checked = False
        End If
    End Sub
    Private Sub HideLabelsNotConnected()
        Dim allLabels() As Label = {Label2,
            Label3, Label4, Label5, Label6, Label7, Label9, Label10, Label11, Label12, Label13, Label14, Label14, Label15,
            Label16, Label17, Label18, Label19, Label20}

        For Each label As Label In allLabels
            If label IsNot Nothing AndAlso
                labelToConnect1 IsNot label AndAlso
                labelToConnect2 IsNot label AndAlso
                labelToConnect3 IsNot label AndAlso
                labelToConnect4 IsNot label AndAlso
                labelToConnect5 IsNot label AndAlso
                labelToConnect6 IsNot label AndAlso
                labelToConnect7 IsNot label AndAlso
                labelToConnect8 IsNot label AndAlso
                labelToConnect9 IsNot label AndAlso
                labelToConnect10 IsNot label AndAlso
                labelToConnect14 IsNot label AndAlso
                labelToConnect15 IsNot label AndAlso
                labelToConnect16 IsNot label AndAlso
                labelToConnect17 IsNot label AndAlso
                labelToConnect18 IsNot label AndAlso
                labelToConnect19 IsNot label AndAlso
                labelToConnect20 IsNot label AndAlso
                labelToConnect21 IsNot label AndAlso
                labelToConnect22 IsNot label AndAlso
                labelToConnect23 IsNot label AndAlso
                labelToConnect24 IsNot label AndAlso
                labelToConnect25 IsNot label AndAlso
                labelToConnect26 IsNot label AndAlso
                labelToConnect27 IsNot label Then
                label.Visible = False
            End If
        Next
    End Sub
    Private Async Function ShowLabelsAsync() As Task
        Dim delay As Integer = 100

        Label1.Visible = True
        Await Task.Delay(delay)

        If labelToConnect1 IsNot Nothing Then
            labelToConnect1.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect2 IsNot Nothing Then
            labelToConnect2.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect3 IsNot Nothing Then
            labelToConnect3.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect4 IsNot Nothing Then
            labelToConnect4.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect5 IsNot Nothing Then
            labelToConnect5.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect6 IsNot Nothing Then
            labelToConnect6.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect6 IsNot Nothing Then
            labelToConnect6.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect7 IsNot Nothing Then
            labelToConnect7.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect8 IsNot Nothing Then
            labelToConnect8.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect9 IsNot Nothing Then
            labelToConnect9.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect10 IsNot Nothing Then
            labelToConnect10.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect14 IsNot Nothing Then
            labelToConnect14.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect15 IsNot Nothing Then
            labelToConnect15.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect16 IsNot Nothing Then
            labelToConnect16.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect17 IsNot Nothing Then
            labelToConnect17.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect18 IsNot Nothing Then
            labelToConnect18.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect19 IsNot Nothing Then
            labelToConnect19.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect20 IsNot Nothing Then
            labelToConnect20.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect21 IsNot Nothing Then
            labelToConnect21.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect22 IsNot Nothing Then
            labelToConnect22.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect23 IsNot Nothing Then
            labelToConnect23.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect24 IsNot Nothing Then
            labelToConnect24.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect25 IsNot Nothing Then
            labelToConnect25.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect26 IsNot Nothing Then
            labelToConnect26.Visible = True
            Await Task.Delay(delay)
        End If
        If labelToConnect27 IsNot Nothing Then
            labelToConnect27.Visible = True
            Await Task.Delay(delay)
        End If
    End Function
    Private drawLine As Boolean = False
    Private labelToConnect1 As Label
    Private labelToConnect2 As Label
    Private labelToConnect3 As Label
    Private labelToConnect4 As Label
    Private labelToConnect5 As Label
    Private labelToConnect6 As Label
    Private labelToConnect7 As Label
    Private labelToConnect8 As Label
    Private labelToConnect9 As Label
    Private labelToConnect10 As Label
    Private labelToConnect14 As Label
    Private labelToConnect15 As Label
    Private labelToConnect16 As Label
    Private labelToConnect17 As Label
    Private labelToConnect18 As Label
    Private labelToConnect19 As Label
    Private labelToConnect20 As Label
    Private labelToConnect21 As Label
    Private labelToConnect22 As Label
    Private labelToConnect23 As Label
    Private labelToConnect24 As Label
    Private labelToConnect25 As Label
    Private labelToConnect26 As Label
    Private labelToConnect27 As Label
    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label2
        labelToConnect3 = Nothing
        labelToConnect4 = Nothing
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "62.51"
    End Sub

    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Nothing
        labelToConnect3 = Nothing
        labelToConnect4 = Nothing
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "56.72"
    End Sub

    Private Async Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label4
        labelToConnect3 = Nothing
        labelToConnect4 = Nothing
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "66.58"
    End Sub

    Private Async Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label4
        labelToConnect3 = Label5
        labelToConnect4 = Nothing
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "122.08"
    End Sub

    Private Async Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label4
        labelToConnect3 = Label5
        labelToConnect4 = Label6
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "168.94"
    End Sub
    Private Async Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label4
        labelToConnect3 = Label5
        labelToConnect4 = Label6
        labelToConnect5 = Label7
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Nothing
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "240.29"
    End Sub
    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        If drawLine Then
            ' Get the coordinates of the labels
            Dim label1Location As Point = Label1.Location
            Dim labelCenter1 As Point
            Dim labelCenter2 As Point
            Dim labelCenter3 As Point
            Dim labelCenter4 As Point
            Dim labelCenter5 As Point
            Dim labelCenter6 As Point
            Dim labelCenter7 As Point
            Dim labelCenter8 As Point
            Dim labelCenter9 As Point
            Dim labelCenter10 As Point
            Dim labelCenter14 As Point
            Dim labelCenter15 As Point
            Dim labelCenter16 As Point
            Dim labelCenter17 As Point
            Dim labelCenter18 As Point
            Dim labelCenter19 As Point
            Dim labelCenter20 As Point
            Dim labelCenter21 As Point
            Dim labelCenter22 As Point
            Dim labelCenter23 As Point
            Dim labelCenter24 As Point
            Dim labelCenter25 As Point
            Dim labelCenter26 As Point
            Dim labelCenter27 As Point

            If labelToConnect1 IsNot Nothing Then
                Dim labelLocation1 As Point = labelToConnect1.Location
                labelCenter1 = New Point(labelLocation1.X - PictureBox1.Left + labelToConnect1.Width \ 2, labelLocation1.Y - PictureBox1.Top + labelToConnect1.Height \ 2)
            Else
                labelCenter1 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If

            If labelToConnect2 IsNot Nothing Then
                Dim labelLocation2 As Point = labelToConnect2.Location
                labelCenter2 = New Point(labelLocation2.X - PictureBox1.Left + labelToConnect2.Width \ 2, labelLocation2.Y - PictureBox1.Top + labelToConnect2.Height \ 2)
            Else
                labelCenter2 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If

            If labelToConnect3 IsNot Nothing Then
                Dim labelLocation3 As Point = labelToConnect3.Location
                labelCenter3 = New Point(labelLocation3.X - PictureBox1.Left + labelToConnect3.Width \ 2, labelLocation3.Y - PictureBox1.Top + labelToConnect3.Height \ 2)
            Else
                labelCenter3 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If

            If labelToConnect4 IsNot Nothing Then
                Dim labelLocation4 As Point = labelToConnect4.Location
                labelCenter4 = New Point(labelLocation4.X - PictureBox1.Left + labelToConnect4.Width \ 2, labelLocation4.Y - PictureBox1.Top + labelToConnect4.Height \ 2)
            Else
                labelCenter4 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If

            If labelToConnect5 IsNot Nothing Then
                Dim labelLocation5 As Point = labelToConnect5.Location
                labelCenter5 = New Point(labelLocation5.X - PictureBox1.Left + labelToConnect5.Width \ 2, labelLocation5.Y - PictureBox1.Top + labelToConnect5.Height \ 2)
            Else
                labelCenter5 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect6 IsNot Nothing Then
                Dim labelLocation6 As Point = labelToConnect6.Location
                labelCenter6 = New Point(labelLocation6.X - PictureBox1.Left + labelToConnect6.Width \ 2, labelLocation6.Y - PictureBox1.Top + labelToConnect6.Height \ 2)
            Else
                labelCenter6 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect7 IsNot Nothing Then
                Dim labelLocation7 As Point = labelToConnect7.Location
                labelCenter7 = New Point(labelLocation7.X - PictureBox1.Left + labelToConnect7.Width \ 2, labelLocation7.Y - PictureBox1.Top + labelToConnect7.Height \ 2)
            Else
                labelCenter7 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect8 IsNot Nothing Then
                Dim labelLocation8 As Point = labelToConnect8.Location
                labelCenter8 = New Point(labelLocation8.X - PictureBox1.Left + labelToConnect8.Width \ 2, labelLocation8.Y - PictureBox1.Top + labelToConnect8.Height \ 2)
            Else
                labelCenter8 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect9 IsNot Nothing Then
                Dim labelLocation9 As Point = labelToConnect9.Location
                labelCenter9 = New Point(labelLocation9.X - PictureBox1.Left + labelToConnect9.Width \ 2, labelLocation9.Y - PictureBox1.Top + labelToConnect9.Height \ 2)
            Else
                labelCenter9 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect10 IsNot Nothing Then
                Dim labelLocation10 As Point = labelToConnect10.Location
                labelCenter10 = New Point(labelLocation10.X - PictureBox1.Left + labelToConnect10.Width \ 2, labelLocation10.Y - PictureBox1.Top + labelToConnect10.Height \ 2)
            Else
                labelCenter10 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect14 IsNot Nothing Then
                Dim labelLocation14 As Point = labelToConnect14.Location
                labelCenter14 = New Point(labelLocation14.X - PictureBox1.Left + labelToConnect14.Width \ 2, labelLocation14.Y - PictureBox1.Top + labelToConnect14.Height \ 2)
            Else
                labelCenter14 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect15 IsNot Nothing Then
                Dim labelLocation15 As Point = labelToConnect15.Location
                labelCenter15 = New Point(labelLocation15.X - PictureBox1.Left + labelToConnect15.Width \ 2, labelLocation15.Y - PictureBox1.Top + labelToConnect15.Height \ 2)
            Else
                labelCenter15 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect16 IsNot Nothing Then
                Dim labelLocation16 As Point = labelToConnect16.Location
                labelCenter16 = New Point(labelLocation16.X - PictureBox1.Left + labelToConnect16.Width \ 2, labelLocation16.Y - PictureBox1.Top + labelToConnect16.Height \ 2)
            Else
                labelCenter16 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect17 IsNot Nothing Then
                Dim labelLocation17 As Point = labelToConnect17.Location
                labelCenter17 = New Point(labelLocation17.X - PictureBox1.Left + labelToConnect17.Width \ 2, labelLocation17.Y - PictureBox1.Top + labelToConnect17.Height \ 2)
            Else
                labelCenter17 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect18 IsNot Nothing Then
                Dim labelLocation18 As Point = labelToConnect18.Location
                labelCenter18 = New Point(labelLocation18.X - PictureBox1.Left + labelToConnect18.Width \ 2, labelLocation18.Y - PictureBox1.Top + labelToConnect18.Height \ 2)
            Else
                labelCenter18 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect19 IsNot Nothing Then
                Dim labelLocation19 As Point = labelToConnect19.Location
                labelCenter19 = New Point(labelLocation19.X - PictureBox1.Left + labelToConnect19.Width \ 2, labelLocation19.Y - PictureBox1.Top + labelToConnect19.Height \ 2)
            Else
                labelCenter19 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect20 IsNot Nothing Then
                Dim labelLocation20 As Point = labelToConnect20.Location
                labelCenter20 = New Point(labelLocation20.X - PictureBox1.Left + labelToConnect20.Width \ 2, labelLocation20.Y - PictureBox1.Top + labelToConnect20.Height \ 2)
            Else
                labelCenter20 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect21 IsNot Nothing Then
                Dim labelLocation21 As Point = labelToConnect21.Location
                labelCenter21 = New Point(labelLocation21.X - PictureBox1.Left + labelToConnect21.Width \ 2, labelLocation21.Y - PictureBox1.Top + labelToConnect21.Height \ 2)
            Else
                labelCenter21 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect22 IsNot Nothing Then
                Dim labelLocation22 As Point = labelToConnect22.Location
                labelCenter22 = New Point(labelLocation22.X - PictureBox1.Left + labelToConnect22.Width \ 2, labelLocation22.Y - PictureBox1.Top + labelToConnect22.Height \ 2)
            Else
                labelCenter22 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect23 IsNot Nothing Then
                Dim labelLocation23 As Point = labelToConnect23.Location
                labelCenter23 = New Point(labelLocation23.X - PictureBox1.Left + labelToConnect23.Width \ 2, labelLocation23.Y - PictureBox1.Top + labelToConnect23.Height \ 2)
            Else
                labelCenter23 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect24 IsNot Nothing Then
                Dim labelLocation24 As Point = labelToConnect24.Location
                labelCenter24 = New Point(labelLocation24.X - PictureBox1.Left + labelToConnect24.Width \ 2, labelLocation24.Y - PictureBox1.Top + labelToConnect24.Height \ 2)
            Else
                labelCenter24 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect25 IsNot Nothing Then
                Dim labelLocation25 As Point = labelToConnect25.Location
                labelCenter25 = New Point(labelLocation25.X - PictureBox1.Left + labelToConnect25.Width \ 2, labelLocation25.Y - PictureBox1.Top + labelToConnect25.Height \ 2)
            Else
                labelCenter25 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect26 IsNot Nothing Then
                Dim labelLocation26 As Point = labelToConnect26.Location
                labelCenter26 = New Point(labelLocation26.X - PictureBox1.Left + labelToConnect26.Width \ 2, labelLocation26.Y - PictureBox1.Top + labelToConnect26.Height \ 2)
            Else
                labelCenter26 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If
            If labelToConnect27 IsNot Nothing Then
                Dim labelLocation27 As Point = labelToConnect27.Location
                labelCenter27 = New Point(labelLocation27.X - PictureBox1.Left + labelToConnect27.Width \ 2, labelLocation27.Y - PictureBox1.Top + labelToConnect27.Height \ 2)
            Else
                labelCenter27 = New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height \ 2)
            End If

            ' Create a Pen with a thicker line width (e.g., 2)
            Dim pen As New Pen(Color.Black, 10.5)
            pen.DashStyle = Drawing2D.DashStyle.Dot

            ' Draw lines between the center points of the labels using the thicker pen
            e.Graphics.DrawLine(pen, labelCenter1, New Point(label1Location.X - PictureBox1.Left + Label1.Width \ 2, label1Location.Y - PictureBox1.Top + Label1.Height))

            If labelToConnect2 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter1, labelCenter2)
            End If
            If labelToConnect3 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter2, labelCenter3)
            End If
            If labelToConnect4 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter3, labelCenter4)
            End If
            If labelToConnect5 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter4, labelCenter5)
            End If
            If labelToConnect6 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter5, labelCenter6)
            End If
            If labelToConnect7 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter6, labelCenter7)
            End If
            If labelToConnect8 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter7, labelCenter8)
            End If
            If labelToConnect9 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter3, labelCenter9)
            End If
            If labelToConnect10 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter9, labelCenter10)
            End If
            If labelToConnect14 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter10, labelCenter14)
            End If
            If labelToConnect15 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter14, labelCenter15)
            End If
            If labelToConnect16 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter15, labelCenter16)
            End If
            If labelToConnect17 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter16, labelCenter17)
            End If
            If labelToConnect18 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter2, labelCenter18)
            End If
            If labelToConnect19 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter2, labelCenter19)
            End If
            If labelToConnect20 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter19, labelCenter20)
            End If
            If labelToConnect21 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter20, labelCenter21)
            End If
            If labelToConnect22 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter21, labelCenter22)
            End If
            If labelToConnect23 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter22, labelCenter23)
            End If
            If labelToConnect24 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter23, labelCenter24)
            End If
            If labelToConnect25 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter24, labelCenter25)
            End If
            If labelToConnect26 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter17, labelCenter26)
            End If
            If labelToConnect27 IsNot Nothing Then
                e.Graphics.DrawLine(pen, labelCenter26, labelCenter27)
            End If
            pen.Dispose()
            drawLine = False
        End If
    End Sub

    Private Async Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Label6
            labelToConnect5 = Label7
            labelToConnect6 = Label9
            labelToConnect7 = Label10
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "301.59"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Label16
            labelToConnect17 = Label17
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Label11
            labelToConnect27 = Label10
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "491.8"
        End If
    End Sub

    Private Async Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Label6
            labelToConnect5 = Label7
            labelToConnect6 = Label9
            labelToConnect7 = Label10
            labelToConnect8 = Label11
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "408.98"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Label16
            labelToConnect17 = Label17
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Label11
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "401.25"
        End If
    End Sub

    Private Async Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "196.82"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label12
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "252.12"
        End If
    End Sub

    Private Async Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "235.44"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "249.71"
        End If
    End Sub

    Private Async Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "283.83"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label14
            labelToConnect23 = Label15
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "271.2"
        End If
    End Sub

    Private Async Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Label16
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "334.26"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label14
            labelToConnect23 = Label15
            labelToConnect24 = Label16
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "321.63"
        End If
    End Sub

    Private Async Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Label16
            labelToConnect17 = Label17
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "372.33"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label14
            labelToConnect23 = Label15
            labelToConnect24 = Label16
            labelToConnect25 = Label17
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "359.7"
        End If
    End Sub

    Private Async Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Label16
            labelToConnect17 = Label17
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "380.25"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label14
            labelToConnect23 = Label15
            labelToConnect24 = Label16
            labelToConnect25 = Label17
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "366.76"
        End If
    End Sub

    Private Async Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        drawLine = True
        labelToConnect1 = Label3
        labelToConnect2 = Label2
        labelToConnect3 = Nothing
        labelToConnect4 = Nothing
        labelToConnect5 = Nothing
        labelToConnect6 = Nothing
        labelToConnect7 = Nothing
        labelToConnect8 = Nothing
        labelToConnect9 = Nothing
        labelToConnect10 = Nothing
        labelToConnect14 = Nothing
        labelToConnect15 = Nothing
        labelToConnect16 = Nothing
        labelToConnect17 = Nothing
        labelToConnect18 = Label18
        labelToConnect19 = Nothing
        labelToConnect20 = Nothing
        labelToConnect21 = Nothing
        labelToConnect22 = Nothing
        labelToConnect23 = Nothing
        labelToConnect24 = Nothing
        labelToConnect25 = Nothing
        labelToConnect26 = Nothing
        labelToConnect27 = Nothing
        PictureBox1.Invalidate()
        HideLabelsNotConnected()
        Await ShowLabelsAsync()
        TextBox3.Text = "102.51"
    End Sub

    Private Async Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "186.31"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label20
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "256.42"
        End If
    End Sub

    Private Async Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "186.31"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label20
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "236.42"
        End If
    End Sub

    Private Async Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "161.21"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label20
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "256.42"
        End If
    End Sub


    Private Sub viewer()

        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Chaos Knight\Documents\rooms.accdb"


        Dim query As String = "SELECT * FROM roomm"


        Using connection As New OleDbConnection(connectionString)
            connection.Open()

            Dim adapter As New OleDbDataAdapter(query, connection)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            DataGridView1.DataSource = dt
        End Using
        For Each column As DataGridViewColumn In DataGridView1.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
        DataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            TextBox2.Text = selectedRow.Cells(0).Value.ToString()

            Dim buttonMappings As New Dictionary(Of String, Button)()
            buttonMappings.Add("21 ANDAYA BUILDING", Button5)
            buttonMappings.Add("08 LABORATORY HIGH SCHOOL", Button19)
            buttonMappings.Add("03 ACADEMIC BUILDING", Button9)
            buttonMappings.Add("15 CERAMICS BUILDING 1", Button11)
            buttonMappings.Add("19 AUTOMOTIVE BUILDING", Button6)
            buttonMappings.Add("13.14 MARITIME BUILDING", Button16)
            buttonMappings.Add("04 STUDENT CENTER", Button17)
            buttonMappings.Add("12 MOTORPOOL BUILDING", Button13)
            buttonMappings.Add("11 SCIENCE AND TECHNOLOGY BUILDING", Button12)
            buttonMappings.Add("18 TECHNOLOGY BUILDING", Button7)

            Dim textBoxValue As String = TextBox2.Text.ToUpper()

            If buttonMappings.ContainsKey(textBoxValue) Then
                Dim targetButton As Button = buttonMappings(textBoxValue)
                targetButton.PerformClick()
            Else
                MessageBox.Show("No matches found. Please use map legend and manually check the building location.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim searchTerm As String = TextBox1.Text.Trim().ToLower()
        If String.IsNullOrWhiteSpace(searchTerm) Then
            MessageBox.Show("Field is empty. Type at least one character", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
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
            MessageBox.Show("No matches found. Please use the map legend and manually check the building location.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Private Async Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label14
            labelToConnect15 = Label15
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "273.83"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Label20
            labelToConnect21 = Label13
            labelToConnect22 = Label14
            labelToConnect23 = Label15
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "261.2"
        End If
    End Sub

    Private Async Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If CheckBox1.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label2
            labelToConnect3 = Nothing
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Nothing
            labelToConnect10 = Nothing
            labelToConnect14 = Nothing
            labelToConnect15 = Nothing
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Label18
            labelToConnect19 = Label19
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "151.21"
        End If
        If CheckBox2.Checked Then
            drawLine = True
            labelToConnect1 = Label3
            labelToConnect2 = Label4
            labelToConnect3 = Label5
            labelToConnect4 = Nothing
            labelToConnect5 = Nothing
            labelToConnect6 = Nothing
            labelToConnect7 = Nothing
            labelToConnect8 = Nothing
            labelToConnect9 = Label12
            labelToConnect10 = Label13
            labelToConnect14 = Label20
            labelToConnect15 = Label19
            labelToConnect16 = Nothing
            labelToConnect17 = Nothing
            labelToConnect18 = Nothing
            labelToConnect19 = Nothing
            labelToConnect20 = Nothing
            labelToConnect21 = Nothing
            labelToConnect22 = Nothing
            labelToConnect23 = Nothing
            labelToConnect24 = Nothing
            labelToConnect25 = Nothing
            labelToConnect26 = Nothing
            labelToConnect27 = Nothing
            PictureBox1.Invalidate()
            HideLabelsNotConnected()
            Await ShowLabelsAsync()
            TextBox3.Text = "266.42"
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Form18.Show()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\techbuild.jpg")
        form17.Show()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\automotive.jpg")
        form17.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\maritime.jpg")
        form17.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\annex.jpg")
        form17.Show()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\motorpool.jpg")
        form17.Show()
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\scitech.jpg")
        form17.Show()
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\postgrad.jpg")
        form17.Show()
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\cer1.jpg")
        form17.Show()
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\cer2.jpg")
        form17.Show()
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\cer3.jpg")
        form17.Show()
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\fitness.jpg")
        form17.Show()
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\lhs.jpg")
        form17.Show()
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\dorm.jpg")
        form17.Show()
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\hostel.jpg")
        form17.Show()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\library.jpg")
        form17.Show()
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\studCent.jpg")
        form17.Show()
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\nursing.jpg")
        form17.Show()
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\acad2.jpg")
        form17.Show()
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\acad.jpg")
        form17.Show()
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\gym.jpg")
        form17.Show()
    End Sub

    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\andaya.jpg")
        form17.Show()
    End Sub

    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\stage.jpg")
        form17.Show()
    End Sub

    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        Dim form17 As New Form17()
        form17.PictureBox1.Image = Image.FromFile("D:\Buildings\court.jpg")
        form17.Show()
    End Sub

    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles Button46.Click
        Form6.ShowDialog()
    End Sub
End Class