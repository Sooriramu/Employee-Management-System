Imports System.Data.SqlClient
Imports System.Data

Public Class Form3

    Dim con2 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim command As New SqlCommand
    Dim n1 As String
    Dim n2 As Integer
    Dim n3 As Integer
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con2.Open()
        DateTimePicker2.Value = Date.Now
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        cb1.SelectedIndex = 0
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" Then
            If TextBox3.Text.Length = 10 Then

                n2 = Format(DateTimePicker1.Value, "yyyy")
                n3 = Format(DateTimePicker2.Value, "yyyy")
                n3 = n3 - n2
                If n3 >= 18 Then
                    If RichTextBox1.Text <> "" Then
                        If TextBox7.Text <> "" Then

                            TextBox1.Text = StrConv(TextBox1.Text, vbProperCase)
                            TextBox6.Text = TextBox1.Text + " " + TextBox5.Text

                            Try

                                If cb1.Text = "Men mazdoor" Then

                                    Dim cmd3c As New SqlCommand("select mm FROM login where username='psg' ", con2)
                                    Dim dr As SqlDataReader
                                    dr = cmd3c.ExecuteReader
                                    dr.Read()
                                    n1 = dr.GetValue(0)
                                    n1 = n1 + 1
                                    dr.Close()
                                    TextBox4.Text = n1
                                    TextBox2.Text = "mm" + n1

                                    Dim cmd3a As New SqlCommand("insert into amm(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox6.Text + "')  ", con2)
                                    cmd3a.ExecuteNonQuery()

                                    Dim cmd3d As New SqlCommand("insert into otltmm (Name,ID) values('" + TextBox6.Text + "','" + TextBox2.Text + "')", con2)
                                    cmd3d.ExecuteNonQuery()

                                    Dim cmd3f As New SqlCommand("update login set mm='" + TextBox4.Text + "' where username='psg'  ", con2)
                                    cmd3f.ExecuteNonQuery()

                                    MessageBox.Show("Added")


                                ElseIf cb1.Text = "Carpenter" Then

                                    Dim cmd3c As New SqlCommand("select c FROM login where username='psg' ", con2)
                                    Dim dr As SqlDataReader
                                    dr = cmd3c.ExecuteReader
                                    dr.Read()
                                    n1 = dr.GetValue(0)
                                    n1 = n1 + 1
                                    dr.Close()
                                    TextBox4.Text = n1
                                    TextBox2.Text = "c" + n1

                                    Dim cmd3a As New SqlCommand("insert into ac(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox6.Text + "')  ", con2)
                                    cmd3a.ExecuteNonQuery()

                                    Dim cmd3d As New SqlCommand("insert into otltc (Name,ID) values('" + TextBox6.Text + "','" + TextBox2.Text + "')", con2)
                                    cmd3d.ExecuteNonQuery()

                                    Dim cmd3f As New SqlCommand("update login set c='" + TextBox4.Text + "' where username='psg'  ", con2)
                                    cmd3f.ExecuteNonQuery()

                                    MessageBox.Show("Added")

                                ElseIf cb1.Text = "Women mazdoor" Then

                                    Dim cmd3c As New SqlCommand("select wm FROM login where username='psg' ", con2)
                                    Dim dr As SqlDataReader
                                    dr = cmd3c.ExecuteReader
                                    dr.Read()
                                    n1 = dr.GetValue(0)
                                    n1 = n1 + 1
                                    dr.Close()
                                    TextBox4.Text = n1
                                    TextBox2.Text = "wm" + n1

                                    Dim cmd3a As New SqlCommand("insert into awm(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox6.Text + "')  ", con2)
                                    cmd3a.ExecuteNonQuery()

                                    Dim cmd3d As New SqlCommand("insert into otltwm (Name,ID) values('" + TextBox6.Text + "','" + TextBox2.Text + "')", con2)
                                    cmd3d.ExecuteNonQuery()

                                    Dim cmd3f As New SqlCommand("update login set wm='" + TextBox4.Text + "' where username='psg'  ", con2)
                                    cmd3f.ExecuteNonQuery()

                                    MessageBox.Show("Added")

                                ElseIf cb1.Text = "Mason" Then

                                    Dim cmd3c As New SqlCommand("select m FROM login where username='psg' ", con2)
                                    Dim dr As SqlDataReader
                                    dr = cmd3c.ExecuteReader
                                    dr.Read()
                                    n1 = dr.GetValue(0)
                                    n1 = n1 + 1
                                    dr.Close()
                                    TextBox4.Text = n1
                                    TextBox2.Text = "m" + n1

                                    Dim cmd3a As New SqlCommand("insert into am(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox6.Text + "')  ", con2)
                                    cmd3a.ExecuteNonQuery()

                                    Dim cmd3d As New SqlCommand("insert into otltm (Name,ID) values('" + TextBox6.Text + "','" + TextBox2.Text + "')", con2)
                                    cmd3d.ExecuteNonQuery()

                                    Dim cmd3f As New SqlCommand("update login set m='" + TextBox4.Text + "' where username='psg'  ", con2)
                                    cmd3f.ExecuteNonQuery()

                                    MessageBox.Show("Added")

                                ElseIf cb1.Text = "Gardener" Then

                                    Dim cmd3c As New SqlCommand("select g FROM login where username='psg' ", con2)
                                    Dim dr As SqlDataReader
                                    dr = cmd3c.ExecuteReader
                                    dr.Read()
                                    n1 = dr.GetValue(0)
                                    n1 = n1 + 1
                                    dr.Close()
                                    TextBox4.Text = n1
                                    TextBox2.Text = "g" + n1

                                    Dim cmd3a As New SqlCommand("insert into ag(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox6.Text + "')  ", con2)
                                    cmd3a.ExecuteNonQuery()

                                    Dim cmd3d As New SqlCommand("insert into otltg (Name,ID) values('" + TextBox6.Text + "','" + TextBox2.Text + "')", con2)
                                    cmd3d.ExecuteNonQuery()

                                    Dim cmd3f As New SqlCommand("update login set g='" + TextBox4.Text + "' where username='psg'  ", con2)
                                    cmd3f.ExecuteNonQuery()

                                    MessageBox.Show("Added")

                                End If

                                Dim cmd3e As New SqlCommand("insert into register values('" + TextBox2.Text + "','" + TextBox6.Text + "','" + TextBox3.Text + "', '" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "','" + RichTextBox1.Text + "','" + cb1.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "','" + TextBox7.Text + "')", con2)
                                cmd3e.ExecuteNonQuery()


                                TextBox1.Text = ""
                                TextBox2.Text = ""
                                TextBox3.Text = ""
                                TextBox4.Text = ""
                                DateTimePicker1.Value = DateTime.Now
                                RichTextBox1.Text = ""

                                Me.Close()
                            Catch ex As Exception
                                MessageBox.Show(ex.ToString)
                            End Try

                        Else
                            MessageBox.Show("Enter the salart amount")

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
            MessageBox.Show("Enter name")
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        TextBox1.Text = ""
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
    End Sub

    Private Sub Form3_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        Me.Hide()
        Form1.Show()

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

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            If TextBox1.Text = "" Then

                Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890\| "
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If


            Else
                Dim dac As String = ",<>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890\|"
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox1.KeyPress
        Try
            If RichTextBox1.Text = "" Then
                Dim dac As String = " "
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Enter the address")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress

        Try
            If TextBox5.Text = "" Then

                Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890\| "
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If


            Else
                Dim dac As String = ",<>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890\|"
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Hide()
        Form11.Show()

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