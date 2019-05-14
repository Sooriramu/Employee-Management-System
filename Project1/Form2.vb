Imports System.Data
Imports System.Data.SqlClient
Public Class Form2
    Dim con1 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim n2, n3, n4, n5
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button3.Hide()
        Label3.Hide()
        con1.Open()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "psg" Then

            Dim command As New SqlCommand("select * from login where Username = '" + TextBox1.Text + "' and Password = '" + TextBox2.Text + "'", con1)
            command.Parameters.Add("@username", SqlDbType.VarChar).Value = TextBox1.Text
            command.Parameters.Add("@password", SqlDbType.VarChar).Value = TextBox2.Text
            Dim adapter As New SqlDataAdapter(command)
            Dim table As New DataTable()
            adapter.Fill(table)
            If table.Rows.Count() <= 0 Then
                MessageBox.Show("Username Or Password Are Invalid")
                TextBox1.Text = ""
                TextBox2.Text = ""
            Else
                MessageBox.Show("Login Successful")
                Dim obj As New Form1
                obj.str = TextBox1.Text
                obj.Show()
                TextBox1.Text = ""
                TextBox2.Text = ""
                Me.Hide()

            End If
        Else
            MessageBox.Show("Username is invalid")
            TextBox1.Text = ""
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Dim cmd2d As New SqlCommand("select pin from login where username='psg'", con1)
        Dim dr1 As SqlDataReader
        dr1 = cmd2d.ExecuteReader
        dr1.Read()
        n2 = dr1.GetValue(0)
        dr1.Close()
        n3 = InputBox("Enter pin", "Security check")
        If n2 = n3 Then
            Form4.Show()
        Else
            MessageBox.Show("Incorrect pin")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button1.Visible = True Then
            Dim cmd2b As New SqlCommand("select pin from login where username='psg'", con1)
            Dim dr As SqlDataReader
            dr = cmd2b.ExecuteReader
            dr.Read()
            n2 = dr.GetValue(0)
            dr.Close()
            n3 = InputBox("Enter pin", "Security check")
            If n2 = n3 Then
                TextBox1.Text = ""
                TextBox2.Text = ""
                Button1.Hide()
                Button3.Show()
                Label3.Show()
            Else
                MessageBox.Show("Invalid pin")
            End If

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim cmd2c As New SqlCommand("select username from login ", con1)
            Dim dr1 As SqlDataReader
            dr1 = cmd2c.ExecuteReader
            n5 = TextBox1.Text
            While dr1.Read
                n4 = dr1.GetString(0)

                If n4 = n5 Then
                    MessageBox.Show("Username already exists")

                End If
            End While
            dr1.Close()
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            Try
                Dim cmd2a As New SqlCommand("insert into login (username,password) values('" + TextBox1.Text + "','" + TextBox2.Text + "') ", con1)
                cmd2a.ExecuteNonQuery()
                MessageBox.Show("Registered successfully")

            Catch ex1 As Exception
                MessageBox.Show(ex1.Message)
            End Try
            TextBox1.Text = ""
            TextBox2.Text = ""
            Button3.Hide()
            Label3.Hide()
            Button1.Show()
        Else
            MessageBox.Show("Textfields cannot be empty")
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Button3.Hide()
        Label3.Hide()
        TextBox1.Text = ""
        TextBox2.Text = ""
        Button1.Show()
    End Sub
End Class
