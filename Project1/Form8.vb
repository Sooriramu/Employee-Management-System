Imports System.Data.SqlClient
Public Class Form8
    Dim con8 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim a, b, c, d, e1
    Dim s1

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim officeexcel As New Microsoft.Office.Interop.Excel.Application
        officeexcel = CreateObject("Excel.application")
        Dim workbook As Object = officeexcel.Workbooks.Add("C:\Custom Office Templates\proj12.xltx")
        officeexcel.Visible = True

        Try

            With officeexcel
                .Range("D6").Value = Format(DateTimePicker1.Value, "MMM dd, yyyy")
                .Range("E8".ToString).Value = Label10.Text
                .Range("E10".ToString).Value = Label11.Text
                .Range("E12".ToString).Value = Label12.Text
                .Range("E14".ToString).Value = Label13.Text
                .Range("E16".ToString).Value = Label14.Text

            End With

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con8)
            Dim dr As SqlDataReader
            dr = cmd8a.ExecuteReader
            dr.Read()
            a = dr.GetDouble(0)
            Label10.Text = a
            Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")
            dr.Close()

            Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con8)
            Dim dr1 As SqlDataReader
            dr1 = cmd8b.ExecuteReader
            dr1.Read()
            b = dr1.GetDouble(0)
            Label11.Text = b
            Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")
            dr1.Close()

            Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con8)
            Dim dr2 As SqlDataReader
            dr2 = cmd8c.ExecuteReader
            dr2.Read()
            c = dr2.GetDouble(0)
            Label12.Text = c
            Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")
            dr2.Close()

            Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con8)
            Dim dr3 As SqlDataReader
            dr3 = cmd8d.ExecuteReader
            dr3.Read()
            d = dr3.GetDouble(0)
            Label13.Text = d
            Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")
            dr3.Close()

            Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con8)
            Dim dr4 As SqlDataReader
            dr4 = cmd8e.ExecuteReader
            dr4.Read()
            e1 = dr4.GetDouble(0)
            Label14.Text = e1
            Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")
            dr4.Close()

            Label15.Text = a + b + c + d + e1


        Catch ex As Exception
            MessageBox.Show("No data found for this date")
        End Try
    End Sub


    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con8.Open()

        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        DateTimePicker1.Value = DateTime.Now
        Label9.Text = DateTimePicker1.Value.Date.ToString("dd-MM-yyyy")

        Try
            Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con8)
            Dim dr As SqlDataReader
            dr = cmd8a.ExecuteReader
            dr.Read()
            a = dr.GetDouble(0)
            Label10.Text = a
            dr.Close()

            Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con8)
            Dim dr1 As SqlDataReader
            dr1 = cmd8b.ExecuteReader
            dr1.Read()
            b = dr1.GetDouble(0)
            Label11.Text = b
            dr1.Close()

            Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con8)
            Dim dr2 As SqlDataReader
            dr2 = cmd8c.ExecuteReader
            dr2.Read()
            c = dr2.GetDouble(0)
            Label12.Text = c
            dr2.Close()

            Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con8)
            Dim dr3 As SqlDataReader
            dr3 = cmd8d.ExecuteReader
            dr3.Read()
            d = dr3.GetDouble(0)
            Label13.Text = d
            dr3.Close()

            Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con8)
            Dim dr4 As SqlDataReader
            dr4 = cmd8e.ExecuteReader
            dr4.Read()
            e1 = dr4.GetDouble(0)
            Label14.Text = e1
            dr4.Close()

            Label15.Text = a + b + c + d + e1

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Form8_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub
End Class