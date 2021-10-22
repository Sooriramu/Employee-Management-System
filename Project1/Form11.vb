Imports System.Data.SqlClient
Imports System.Globalization

Public Class Form11
    Dim con2 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim command As New SqlCommand
    Dim n2, n3
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If cb2.Text <> "" Then
            If TextBox3.Text.Length = 10 Then
                n2 = Format(DateTimePicker1.Value, "yyyy")
                n3 = Format(DateTimePicker2.Value, "yyyy")
                n3 = n3 - n2
                If n3 >= 18 Then
                    If RichTextBox1.Text <> "" Then
                        If TextBox7.Text <> "" Then

                            Try

                                Dim cmd3a As New SqlCommand("update register set Phno='" + TextBox3.Text + "' ,dob='" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "',address='" + RichTextBox1.Text + "',salary='" + TextBox7.Text + "' where id='" + cb2.Text + "' ", con2)
                                cmd3a.ExecuteNonQuery()

                                MessageBox.Show("Updated successfully")
                                Label7.Text = ""
                                TextBox3.Text = ""
                                DateTimePicker1.Value = Date.Now
                                RichTextBox1.Text = ""
                                TextBox7.Text = ""
                                Call cb1_SelectedIndexChanged(sender, e)

                            Catch ex As Exception
                                MessageBox.Show(ex.ToString)
                            End Try

                        Else
                            MessageBox.Show("Enter the salary amount")
                        End If
                    Else
                        MessageBox.Show("Enter the address")
                    End If
                Else
                    MessageBox.Show("Age below 18 cannot be employed")
                End If
            Else
                MessageBox.Show("Phone number must contain 10 numbers")
            End If
        Else
            MessageBox.Show("Select the id")
        End If

    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        con2.Open()

        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        cb1.SelectedIndex = 0
        Label7.Text = ""
        DateTimePicker2.Value = Date.Now
    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged

        Label7.Text = ""
        TextBox3.Text = ""
        DateTimePicker1.Value = Date.Now
        RichTextBox1.Text = ""

        Dim cmd3b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select id,name from amm order by len(id), id"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select id,name from ac order by len(id), id"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select id,name from awm order by len(id), id"
            ElseIf cb1.Text = "Mason" Then
                query = "select id,name from am order by len(id), id"
            ElseIf cb1.Text = "Gardener" Then
                query = "select id,name from ag order by len(id), id"
            End If
            command = New SqlCommand(query, con2)
            cmd3b.SelectCommand = command
            cmd3b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd3b.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        cb2.Items.Clear()
        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from amm", con2)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader1.Close()

            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1d As New SqlCommand("select id from ac", con2)
                Dim reader2 As SqlDataReader
                reader2 = cmd1d.ExecuteReader

                While reader2.Read
                    Dim sid = reader2.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader2.Close()

            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from awm", con2)
                Dim reader3 As SqlDataReader
                reader3 = cmd1d.ExecuteReader

                While reader3.Read
                    Dim sid = reader3.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader3.Close()

            ElseIf cb1.Text = "Mason" Then
                Dim cmd1d As New SqlCommand("select id from am", con2)
                Dim reader4 As SqlDataReader
                reader4 = cmd1d.ExecuteReader

                While reader4.Read
                    Dim sid = reader4.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader4.Close()

            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1d As New SqlCommand("select id from ag", con2)
                Dim reader5 As SqlDataReader
                reader5 = cmd1d.ExecuteReader

                While reader5.Read
                    Dim sid = reader5.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader5.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub cb2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb2.SelectedIndexChanged

        Try

            Dim cmd1y As New SqlCommand("select phno from register where id='" + cb2.Text + "'", con2)
            Dim dr1 As SqlDataReader
            dr1 = cmd1y.ExecuteReader
            dr1.Read()
            TextBox3.Text = dr1.GetString(0)
            dr1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Try

            Dim cmd1y As New SqlCommand("select dob from register where id='" + cb2.Text + "'", con2)
            Dim dr1 As SqlDataReader
            dr1 = cmd1y.ExecuteReader
            dr1.Read()

            DateTimePicker1.Value = DateTime.ParseExact(dr1.GetValue(0), "dd-MM-yyyy", CultureInfo.InvariantCulture)
            '(dr1.GetValue(0))

            dr1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Try

            Dim cmd1y As New SqlCommand("select address from register where id='" + cb2.Text + "'", con2)
            Dim dr1 As SqlDataReader
            dr1 = cmd1y.ExecuteReader
            dr1.Read()
            RichTextBox1.Text = dr1.GetValue(0)
            dr1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Try

            Dim cmd1y As New SqlCommand("select name from register where id='" + cb2.Text + "'", con2)
            Dim dr1 As SqlDataReader
            dr1 = cmd1y.ExecuteReader
            dr1.Read()
            Label7.Text = dr1.GetValue(0)
            dr1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Try

            Dim cmd1y As New SqlCommand("select salary from register where id='" + cb2.Text + "'", con2)
            Dim dr1 As SqlDataReader
            dr1 = cmd1y.ExecuteReader
            dr1.Read()
            TextBox7.Text = dr1.GetValue(0)
            dr1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Form11_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        Try
            Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`\| "

            If Char.IsLetter(e.KeyChar) = True Then
                e.Handled = True
                MessageBox.Show("Only numbers are allowed")

            ElseIf dac.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
                MessageBox.Show("Only numbers are allowed")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        Try
            Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`\| "

            If Char.IsLetter(e.KeyChar) = True Then
                e.Handled = True
                MessageBox.Show("Only numbers are allowed")

            ElseIf dac.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
                MessageBox.Show("Only numbers are allowed")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class