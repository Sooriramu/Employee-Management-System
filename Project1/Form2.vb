Imports System.Data.SqlClient
Public Class Form2
    Dim con1 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim n2, n3, n4, n5
    Dim n6 = 0
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)

        Button3.Hide()
        Label3.Hide()
        con1.Open()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "psg" Then
            Try
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


                    Form1.Label4.Text = TextBox1.Text

                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    Form1.Show()
                    Me.Hide()


                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            MessageBox.Show("Username is invalid")
            TextBox1.Text = ""
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Try
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
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
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
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        n6 = 0
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then

            Dim cmd2c As New SqlCommand("select username from login ", con1)
            Dim dr1 As SqlDataReader
            dr1 = cmd2c.ExecuteReader
            n5 = TextBox1.Text
            While dr1.Read
                n4 = dr1.GetString(0)

                If n4 = n5 Then
                    MessageBox.Show("Username already exists")
                    n6 = 1
                End If
            End While
            dr1.Close()

            If n6 <> 1 Then
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

            End If
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form2_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Application.Exit()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        Try
            Dim dac As String = " "
            If dac.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
                MessageBox.Show("Blankspaces are not allowed")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

        Try
            Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`"" "
            If dac.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
                MessageBox.Show("Only alphabets and numbers are allowed")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class
