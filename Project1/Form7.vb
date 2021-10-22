Imports System.Data
Imports System.Data.SqlClient

Public Class Form7
    Dim con7 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim m1 As Integer
    Dim m2, b1, b2, b3, b4, b5, b6, b7, j As Integer
    Dim s1, s2, s3, s4, s5, s6, s7, n1, n2, n3, n4, n5, n6, n7, n8, total1, total
    Dim c As Integer
    Dim a1, a2, a3, a4, a5, a6, a7 As Integer
    Dim i As String
    Dim j1,total2
    Dim command As SqlCommand


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        DataGridView1.DataSource = ""

    End Sub
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        con7.Open()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        DateTimePicker3.Value = Date.Now
        DateTimePicker4.Value = Date.Now
        DateTimePicker5.Value = Date.Now
        DateTimePicker6.Value = Date.Now
        DateTimePicker7.Value = Date.Now
        DateTimePicker8.Value = Date.Now
        cb1.SelectedIndex = 0


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If DateTimePicker1.Value <= Date.Now And DateTimePicker2.Value <= Date.Now Then

            If DateTimePicker1.Value <= DateTimePicker2.Value Then

                m1 = 1
                DateTimePicker8.Value = DateTimePicker1.Value.ToShortDateString


                If DateTimePicker8.Value <> DateTimePicker2.Value.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)

                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)
                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)
                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)
                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)
                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)
                End If
                If DateTimePicker8.Value <> DateTimePicker2.Value.Date.ToShortDateString Then
                    m1 = m1 + 1
                    DateTimePicker8.Value = DateTimePicker8.Value.AddDays(1)

                End If

                If m1 = 1 Then

                    DataGridView1.DataSource = ""

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from  ag"
                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    Try
                        Dim cmd7g As New SqlCommand("delete from salary", con7)
                        cmd7g.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    Try

                        If cb1.Text = "Men mazdoor" Then

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name,Working_days) select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]  from amm ", con7)
                            cmd7f.ExecuteNonQuery()


                            Dim cmd7k As New SqlCommand("select count(*) from otltmm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from otltmm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            i = 0
                            While i < n4

                                cb2.SelectedIndex = i
                                i = i + 1
                                total2 = 0
                                total1 = 0
                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = a1
                                Else
                                    total1 = a1
                                End If

                                Dim cmd7zb As New SqlCommand("update salary set OT_hours ='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7zb.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While


                        ElseIf cb1.Text = "Carpenter" Then

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name,Working_days) select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from otltc ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from otltc", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            i = 0
                            While i < n4

                                cb2.SelectedIndex = i
                                i = i + 1
                                total2 = 0
                                total1 = 0
                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = a1
                                Else
                                    total1 = a1
                                End If

                                Dim cmd7zb As New SqlCommand("update salary set OT_hours ='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7zb.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        ElseIf cb1.Text = "Women mazdoor" Then

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name,Working_days) select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from otltwm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from otltwm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            i = 0
                            While i < n4

                                cb2.SelectedIndex = i
                                i = i + 1
                                total2 = 0
                                total1 = 0
                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = a1
                                Else
                                    total1 = a1
                                End If

                                Dim cmd7zb As New SqlCommand("update salary set OT_hours ='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7zb.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        ElseIf cb1.Text = "Mason" Then

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name,Working_days) select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from otltm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from otltm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            i = 0
                            While i < n4

                                cb2.SelectedIndex = i
                                i = i + 1
                                total2 = 0
                                total1 = 0
                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = a1
                                Else
                                    total1 = a1
                                End If

                                Dim cmd7zb As New SqlCommand("update salary set OT_hours ='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7zb.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        ElseIf cb1.Text = "Gardener" Then

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name,Working_days) select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from otltg ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from otltg", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            i = 0
                            While i < n4

                                cb2.SelectedIndex = i
                                i = i + 1
                                total2 = 0
                                total1 = 0
                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg  where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = a1
                                Else
                                    total1 = a1
                                End If

                                Dim cmd7zb As New SqlCommand("update salary set OT_hours ='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7zb.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try


                ElseIf m1 = 2 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from  ag"

                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()


                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1


                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                total = n1 + n2

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7y As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7y.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                total1 = a1 + a2

                                Dim cmd7z As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7z.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                total = n1 + n2

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7y As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7y.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                total1 = a1 + a2

                                Dim cmd7z As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7z.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                total = n1 + n2

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7y As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7y.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                total1 = a1 + a2

                                Dim cmd7z As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7z.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                total = n1 + n2

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7y As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7y.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                total1 = a1 + a2

                                Dim cmd7z As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7z.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                total = n1 + n2

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7x As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7x.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7y As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7y.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                total1 = a1 + a2

                                Dim cmd7z As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7z.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                ElseIf m1 = 3 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from  ag"
                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()

                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                total = n1 + n2 + n3

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7aa.ExecuteReader
                                dr4.Read()
                                a1 = dr4.GetValue(0)
                                dr4.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7ab.ExecuteReader
                                dr5.Read()
                                a2 = dr5.GetValue(0)
                                dr5.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7ac.ExecuteReader
                                dr6.Read()
                                a3 = dr6.GetValue(0)
                                dr6.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                total1 = a1 + a2 + a3

                                Dim cmd7az As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7az.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                total = n1 + n2 + n3

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7aa.ExecuteReader
                                dr4.Read()
                                a1 = dr4.GetValue(0)
                                dr4.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7ab.ExecuteReader
                                dr5.Read()
                                a2 = dr5.GetValue(0)
                                dr5.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7ac.ExecuteReader
                                dr6.Read()
                                a3 = dr6.GetValue(0)
                                dr6.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                total1 = a1 + a2 + a3

                                Dim cmd7az As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7az.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                total = n1 + n2 + n3

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7aa.ExecuteReader
                                dr4.Read()
                                a1 = dr4.GetValue(0)
                                dr4.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7ab.ExecuteReader
                                dr5.Read()
                                a2 = dr5.GetValue(0)
                                dr5.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7ac.ExecuteReader
                                dr6.Read()
                                a3 = dr6.GetValue(0)
                                dr6.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                total1 = a1 + a2 + a3

                                Dim cmd7az As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7az.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then


                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()



                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                total = n1 + n2 + n3

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7aa.ExecuteReader
                                dr4.Read()
                                a1 = dr4.GetValue(0)
                                dr4.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7ab.ExecuteReader
                                dr5.Read()
                                a2 = dr5.GetValue(0)
                                dr5.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7ac.ExecuteReader
                                dr6.Read()
                                a3 = dr6.GetValue(0)
                                dr6.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If


                                total1 = a1 + a2 + a3

                                Dim cmd7az As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7az.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While


                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                total = n1 + n2 + n3

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7aa.ExecuteReader
                                dr4.Read()
                                a1 = dr4.GetValue(0)
                                dr4.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7ab.ExecuteReader
                                dr5.Read()
                                a2 = dr5.GetValue(0)
                                dr5.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7ac.ExecuteReader
                                dr6.Read()
                                a3 = dr6.GetValue(0)
                                dr6.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                total1 = a1 + a2 + a3

                                Dim cmd7az As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7az.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                ElseIf m1 = 4 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "]  from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "]  from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "]  from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "]  from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "]  from  ag"
                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try


                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()

                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                total = n1 + n2 + n3 + n5

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a5 = dr9.GetValue(0)
                                dr9.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try



                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                total = n1 + n2 + n3 + n5

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a5 = dr9.GetValue(0)
                                dr9.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                total = n1 + n2 + n3 + n5

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a5 = dr9.GetValue(0)
                                dr9.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then


                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                total = n1 + n2 + n3 + n5

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a5 = dr9.GetValue(0)
                                dr9.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                total = n1 + n2 + n3 + n5

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a5 = dr9.GetValue(0)
                                dr9.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                ElseIf m1 = 5 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from  ag"

                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try


                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()

                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                total = n1 + n2 + n3 + n5 + n6

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try



                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                total = n1 + n2 + n3 + n5 + n6

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                total = n1 + n2 + n3 + n5 + n6

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then


                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                total = n1 + n2 + n3 + n5 + n6

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                total = n1 + n2 + n3 + n5 + n6

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7aa.ExecuteReader
                                dr6.Read()
                                a1 = dr6.GetValue(0)
                                dr6.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                ElseIf m1 = 6 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)
                    DateTimePicker7.Value = DateTimePicker1.Value.AddDays(5)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select id,name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from  ag"

                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()

                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 = a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try



                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()
                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then


                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 = a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7ab.ExecuteReader
                                dr7.Read()
                                a2 = dr7.GetValue(0)
                                dr7.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                ElseIf m1 = 7 Then
                    DataGridView1.DataSource = ""
                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)
                    DateTimePicker7.Value = DateTimePicker1.Value.AddDays(5)
                    DateTimePicker8.Value = DateTimePicker1.Value.AddDays(6)

                    Dim cmd7a As New SqlDataAdapter
                    Dim dt As New DataTable
                    Dim query As String
                    Dim bsource As New BindingSource
                    Try
                        If cb1.Text = "Men mazdoor" Then
                            query = "select ID,Name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from  amm"
                        ElseIf cb1.Text = "Carpenter" Then
                            query = "select ID,Name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from  ac"
                        ElseIf cb1.Text = "Women mazdoor" Then
                            query = "select ID,Name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from  awm"
                        ElseIf cb1.Text = "Mason" Then
                            query = "select ID,Name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from  am"
                        ElseIf cb1.Text = "Gardener" Then
                            query = "select ID,Name,[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "],[" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from  ag"

                        End If
                        command = New SqlCommand(query, con7)
                        cmd7a.SelectCommand = command
                        cmd7a.Fill(dt)
                        bsource.DataSource = dt
                        DataGridView1.DataSource = bsource
                        cmd7a.Update(dt)

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try


                    Dim cmd7g As New SqlCommand("delete  from salary", con7)
                    cmd7g.ExecuteNonQuery()

                    If cb1.Text = "Men mazdoor" Then

                        i = 0
                        total = 0


                        Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from amm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from amm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from amm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                Dim cmd7p As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7p.ExecuteReader
                                dr7.Read()
                                n8 = dr7.GetDouble(0)
                                dr7.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7 + n8

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr13 As SqlDataReader
                                dr13 = cmd7ab.ExecuteReader
                                dr13.Read()
                                a2 = dr13.GetValue(0)
                                dr13.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                Dim cmd7ah As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con7)
                                Dim dr14 As SqlDataReader
                                dr14 = cmd7ah.ExecuteReader
                                dr14.Read()
                                a7 = dr14.GetValue(0)
                                dr14.Close()
                                If a7 < 0 Then
                                    total2 = total2 + a7
                                    a7 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6 + a7

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While





                    ElseIf cb1.Text = "Carpenter" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ac", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ac ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ac", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                Dim cmd7p As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7p.ExecuteReader
                                dr7.Read()
                                n8 = dr7.GetDouble(0)
                                dr7.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7 + n8

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr13 As SqlDataReader
                                dr13 = cmd7ab.ExecuteReader
                                dr13.Read()
                                a2 = dr13.GetValue(0)
                                dr13.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                Dim cmd7ah As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con7)
                                Dim dr14 As SqlDataReader
                                dr14 = cmd7ah.ExecuteReader
                                dr14.Read()
                                a7 = dr14.GetValue(0)
                                dr14.Close()
                                If a7 < 0 Then
                                    total2 = total2 + a7
                                    a7 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6 + a7

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Women mazdoor" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from awm", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from awm ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from awm", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                Dim cmd7p As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7p.ExecuteReader
                                dr7.Read()
                                n8 = dr7.GetDouble(0)
                                dr7.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7 + n8

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr13 As SqlDataReader
                                dr13 = cmd7ab.ExecuteReader
                                dr13.Read()
                                a2 = dr13.GetValue(0)
                                dr13.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                Dim cmd7ah As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con7)
                                Dim dr14 As SqlDataReader
                                dr14 = cmd7ah.ExecuteReader
                                dr14.Read()
                                a7 = dr14.GetValue(0)
                                dr14.Close()
                                If a7 < 0 Then
                                    total2 = total2 + a7
                                    a7 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6 + a7

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Mason" Then


                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from am", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from am ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from am", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                Dim cmd7p As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7p.ExecuteReader
                                dr7.Read()
                                n8 = dr7.GetDouble(0)
                                dr7.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7 + n8

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr13 As SqlDataReader
                                dr13 = cmd7ab.ExecuteReader
                                dr13.Read()
                                a2 = dr13.GetValue(0)
                                dr13.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                Dim cmd7ah As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con7)
                                Dim dr14 As SqlDataReader
                                dr14 = cmd7ah.ExecuteReader
                                dr14.Read()
                                a7 = dr14.GetValue(0)
                                dr14.Close()
                                If a7 < 0 Then
                                    total2 = total2 + a7
                                    a7 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6 + a7

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    ElseIf cb1.Text = "Gardener" Then

                        i = 0
                        total = 0
                        Try

                            Dim cmd7f As New SqlCommand("insert into salary (ID,Name) select id,name  from ag", con7)
                            cmd7f.ExecuteNonQuery()

                            Dim cmd7k As New SqlCommand("select count(*) from ag ", con7)
                            Dim dr2 As SqlDataReader
                            dr2 = cmd7k.ExecuteReader
                            dr2.Read()
                            n4 = dr2.GetValue(0)
                            dr2.Close()

                            Dim cmd1d As New SqlCommand("select id from ag", con7)
                            Dim reader1 As SqlDataReader
                            reader1 = cmd1d.ExecuteReader
                            cb2.Items.Clear()

                            While reader1.Read
                                Dim sid = reader1.GetString(0)
                                cb2.Items.Add(sid)
                            End While
                            reader1.Close()

                            While i < n4
                                i = i + 1
                                cb2.SelectedIndex = i - 1

                                Dim cmd7h As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr As SqlDataReader
                                dr = cmd7h.ExecuteReader
                                dr.Read()
                                n1 = dr.GetDouble(0)
                                dr.Close()

                                Dim cmd7i As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr1 As SqlDataReader
                                dr1 = cmd7i.ExecuteReader
                                dr1.Read()
                                n2 = dr1.GetDouble(0)
                                dr1.Close()

                                Dim cmd7l As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr3 As SqlDataReader
                                dr3 = cmd7l.ExecuteReader
                                dr3.Read()
                                n3 = dr3.GetDouble(0)
                                dr3.Close()

                                Dim cmd7m As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr4 As SqlDataReader
                                dr4 = cmd7m.ExecuteReader
                                dr4.Read()
                                n5 = dr4.GetDouble(0)
                                dr4.Close()

                                Dim cmd7n As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr5 As SqlDataReader
                                dr5 = cmd7n.ExecuteReader
                                dr5.Read()
                                n6 = dr5.GetDouble(0)
                                dr5.Close()

                                Dim cmd7o As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr6 As SqlDataReader
                                dr6 = cmd7o.ExecuteReader
                                dr6.Read()
                                n7 = dr6.GetDouble(0)
                                dr6.Close()

                                Dim cmd7p As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con7)
                                Dim dr7 As SqlDataReader
                                dr7 = cmd7p.ExecuteReader
                                dr7.Read()
                                n8 = dr7.GetDouble(0)
                                dr7.Close()

                                total = n1 + n2 + n3 + n5 + n6 + n7 + n8

                                Dim cmd7j As New SqlCommand("update salary set Working_days='" + total.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7j.ExecuteNonQuery()

                                total2 = 0

                                Dim cmd7aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr11 As SqlDataReader
                                dr11 = cmd7aa.ExecuteReader
                                dr11.Read()
                                a1 = dr11.GetValue(0)
                                dr11.Close()
                                If a1 < 0 Then
                                    total2 = total2 + a1
                                    a1 = 0
                                End If

                                Dim cmd7ab As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr13 As SqlDataReader
                                dr13 = cmd7ab.ExecuteReader
                                dr13.Read()
                                a2 = dr13.GetValue(0)
                                dr13.Close()
                                If a2 < 0 Then
                                    total2 = total2 + a2
                                    a2 = 0
                                End If

                                Dim cmd7ac As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr8 As SqlDataReader
                                dr8 = cmd7ac.ExecuteReader
                                dr8.Read()
                                a3 = dr8.GetValue(0)
                                dr8.Close()
                                If a3 < 0 Then
                                    total2 = total2 + a3
                                    a3 = 0
                                End If

                                Dim cmd7ad As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr9 As SqlDataReader
                                dr9 = cmd7ad.ExecuteReader
                                dr9.Read()
                                a4 = dr9.GetValue(0)
                                dr9.Close()
                                If a4 < 0 Then
                                    total2 = total2 + a4
                                    a4 = 0
                                End If

                                Dim cmd7af As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr10 As SqlDataReader
                                dr10 = cmd7af.ExecuteReader
                                dr10.Read()
                                a5 = dr10.GetValue(0)
                                dr10.Close()
                                If a5 < 0 Then
                                    total2 = total2 + a5
                                    a5 = 0
                                End If

                                Dim cmd7ag As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr12 As SqlDataReader
                                dr12 = cmd7ag.ExecuteReader
                                dr12.Read()
                                a6 = dr12.GetValue(0)
                                dr12.Close()
                                If a6 < 0 Then
                                    total2 = total2 + a6
                                    a6 = 0
                                End If

                                Dim cmd7ah As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con7)
                                Dim dr14 As SqlDataReader
                                dr14 = cmd7ah.ExecuteReader
                                dr14.Read()
                                a7 = dr14.GetValue(0)
                                dr14.Close()
                                If a7 < 0 Then
                                    total2 = total2 + a7
                                    a7 = 0
                                End If

                                total1 = a1 + a2 + a3 + a4 + a5 + a6 + a7

                                Dim cmd7ae As New SqlCommand("update salary set OT_hours='" + total1.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7ae.ExecuteNonQuery()

                                Dim cmd7za As New SqlCommand("update salary set OT_amount='" + total2.ToString() + "' where id='" + cb2.Text + "'", con7)
                                cmd7za.ExecuteNonQuery()

                            End While

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If


                Else
                    MessageBox.Show("Can fetch records for maximum of 7 days only")

                End If

            Else
                MessageBox.Show("Incorrect date provided")
            End If

        Else
            MessageBox.Show("Enter valid date")
        End If

        Try

            Dim cmd1da As New SqlCommand("select id from salary ", con7)
            Dim reader1a As SqlDataReader
            reader1a = cmd1da.ExecuteReader
            cb2.Items.Clear()
            While reader1a.Read
                Dim sid = reader1a.GetString(0)
                cb2.Items.Add(sid)
            End While
            reader1a.Close()

            For j = 0 To cb2.Items.Count - 1

                cb2.SelectedIndex = j

                Dim cmdz As New SqlCommand("select salary from register where id='" + cb2.Text + "'", con7)
                Dim dr As SqlDataReader
                dr = cmdz.ExecuteReader
                dr.Read()
                TextBox1.Text = dr.GetString(0)
                dr.Close()

                Dim cmdza As New SqlCommand("update salary set Per_day='" + TextBox1.Text + "' where id='" + cb2.Text + "'", con7)
                cmdza.ExecuteNonQuery()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If DataGridView1.Rows.Count <> 0 Then

            Dim officeexcel As New Microsoft.Office.Interop.Excel.Application
            officeexcel = CreateObject("Excel.application")
            Dim workbook As Object = officeexcel.Workbooks.Add("C:\Custom Office Templates\proj11.xltx")
            officeexcel.Visible = True
            Dim da As New SqlDataAdapter
            Dim ds As New DataSet
            Try
                da = New SqlDataAdapter("select * from salary order by len(id), id ", con7)
                da.Fill(ds, "salary")

                With officeexcel
                    .Range("B10").Value = Format(DateTimePicker1.Value, "MMM dd, yyyy")
                    .Range("F10").Value = Format(DateTimePicker2.Value, "MMM dd, yyyy")
                    .Range("C8").Value = cb1.Text
                End With
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            n2 = 1

            DateTimePicker3.Value = DateTimePicker1.Value.ToShortDateString

            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If
            If DateTimePicker3.Value <> DateTimePicker2.Value.ToShortDateString Then
                n2 = n2 + 1
                DateTimePicker3.Value = DateTimePicker3.Value.AddDays(1)
            End If



            If n2 = 1 Then
                Try

                    If cb1.Text = "Men mazdoor" Then

                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then

                        TextBox1.Text = "otltc"
                        Dim cmd8a As New SqlCommand("Select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "]) from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()
                        With officeexcel
                            .Range("E14".ToString).Value = s1
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        TextBox1.Text = "otltwm"
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()
                        With officeexcel
                            .Range("E14".ToString).Value = s1
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        TextBox1.Text = "otltm"
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()
                        With officeexcel
                            .Range("E14".ToString).Value = s1
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        TextBox1.Text = "otltg"
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()
                        With officeexcel
                            .Range("E14".ToString).Value = s1
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                    End If


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            ElseIf n2 = 2 Then

                Try

                    If cb1.Text = "Men mazdoor" Then

                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker2.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker2.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            ElseIf n2 = 3 Then

                Try

                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)

                    If cb1.Text = "Men mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With


                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            ElseIf n2 = 4 Then

                Try

                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)

                    If cb1.Text = "Men mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            ElseIf n2 = 5 Then

                Try

                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)

                    If cb1.Text = "Men mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetValue(0)
                        dr4.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            ElseIf n2 = 6 Then

                Try

                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)
                    DateTimePicker7.Value = DateTimePicker1.Value.AddDays(5)

                    If cb1.Text = "Men mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            ElseIf n2 = 7 Then

                Try

                    DateTimePicker3.Value = DateTimePicker1.Value.AddDays(1)
                    DateTimePicker4.Value = DateTimePicker1.Value.AddDays(2)
                    DateTimePicker5.Value = DateTimePicker1.Value.AddDays(3)
                    DateTimePicker6.Value = DateTimePicker1.Value.AddDays(4)
                    DateTimePicker7.Value = DateTimePicker1.Value.AddDays(5)
                    DateTimePicker8.Value = DateTimePicker1.Value.AddDays(6)

                    If cb1.Text = "Men mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        Dim cmd8g As New SqlCommand("select sum([" + DateTimePicker8.Value.Date.ToString("dd-MM-yyyy") + "])from amm", con7)
                        Dim dr6 As SqlDataReader
                        dr6 = cmd8g.ExecuteReader
                        dr6.Read()
                        s7 = dr6.GetDouble(0)
                        dr6.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                            .Range("E20".ToString).Value = s7
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                        Dim cmd1ef As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltmm", con7)
                        Dim reader2f As SqlDataReader
                        reader2f = cmd1ef.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2f.Read

                            Dim sid = reader2f.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2f.Close()
                        With officeexcel
                            .Range("G20".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Carpenter" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        Dim cmd8g As New SqlCommand("select sum([" + DateTimePicker8.Value.Date.ToString("dd-MM-yyyy") + "])from ac", con7)
                        Dim dr6 As SqlDataReader
                        dr6 = cmd8g.ExecuteReader
                        dr6.Read()
                        s7 = dr6.GetDouble(0)
                        dr6.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                            .Range("E20".ToString).Value = s7
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                        Dim cmd1ef As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltc", con7)
                        Dim reader2f As SqlDataReader
                        reader2f = cmd1ef.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2f.Read

                            Dim sid = reader2f.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2f.Close()
                        With officeexcel
                            .Range("G20".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Women mazdoor" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        Dim cmd8g As New SqlCommand("select sum([" + DateTimePicker8.Value.Date.ToString("dd-MM-yyyy") + "])from awm", con7)
                        Dim dr6 As SqlDataReader
                        dr6 = cmd8g.ExecuteReader
                        dr6.Read()
                        s7 = dr6.GetDouble(0)
                        dr6.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                            .Range("E20".ToString).Value = s7
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                        Dim cmd1ef As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltwm", con7)
                        Dim reader2f As SqlDataReader
                        reader2f = cmd1ef.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2f.Read

                            Dim sid = reader2f.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2f.Close()
                        With officeexcel
                            .Range("G20".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Mason" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        Dim cmd8g As New SqlCommand("select sum([" + DateTimePicker8.Value.Date.ToString("dd-MM-yyyy") + "])from am", con7)
                        Dim dr6 As SqlDataReader
                        dr6 = cmd8g.ExecuteReader
                        dr6.Read()
                        s7 = dr6.GetDouble(0)
                        dr6.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                            .Range("E20".ToString).Value = s7
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                        Dim cmd1ef As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltm", con7)
                        Dim reader2f As SqlDataReader
                        reader2f = cmd1ef.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2f.Read

                            Dim sid = reader2f.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2f.Close()
                        With officeexcel
                            .Range("G20".ToString).Value = s1
                        End With

                    ElseIf cb1.Text = "Gardener" Then
                        Dim cmd8a As New SqlCommand("select sum([" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr As SqlDataReader
                        dr = cmd8a.ExecuteReader
                        dr.Read()
                        s1 = dr.GetDouble(0)
                        dr.Close()

                        Dim cmd8b As New SqlCommand("select sum([" + DateTimePicker3.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr1 As SqlDataReader
                        dr1 = cmd8b.ExecuteReader
                        dr1.Read()
                        s2 = dr1.GetDouble(0)
                        dr1.Close()

                        Dim cmd8c As New SqlCommand("select sum([" + DateTimePicker4.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr2 As SqlDataReader
                        dr2 = cmd8c.ExecuteReader
                        dr2.Read()
                        s3 = dr2.GetDouble(0)
                        dr2.Close()

                        Dim cmd8d As New SqlCommand("select sum([" + DateTimePicker5.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr3 As SqlDataReader
                        dr3 = cmd8d.ExecuteReader
                        dr3.Read()
                        s4 = dr3.GetDouble(0)
                        dr3.Close()

                        Dim cmd8e As New SqlCommand("select sum([" + DateTimePicker6.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr4 As SqlDataReader
                        dr4 = cmd8e.ExecuteReader
                        dr4.Read()
                        s5 = dr4.GetDouble(0)
                        dr4.Close()

                        Dim cmd8f As New SqlCommand("select sum([" + DateTimePicker7.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr5 As SqlDataReader
                        dr5 = cmd8f.ExecuteReader
                        dr5.Read()
                        s6 = dr5.GetDouble(0)
                        dr5.Close()

                        Dim cmd8g As New SqlCommand("select sum([" + DateTimePicker8.Value.Date.ToString("dd-MM-yyyy") + "])from ag", con7)
                        Dim dr6 As SqlDataReader
                        dr6 = cmd8g.ExecuteReader
                        dr6.Read()
                        s7 = dr6.GetDouble(0)
                        dr6.Close()

                        With officeexcel
                            .Range("E14".ToString).Value = s1
                            .Range("E15".ToString).Value = s2
                            .Range("E16".ToString).Value = s3
                            .Range("E17".ToString).Value = s4
                            .Range("E18".ToString).Value = s5
                            .Range("E19".ToString).Value = s6
                            .Range("E20".ToString).Value = s7
                        End With

                        Dim cmd1e As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2 As SqlDataReader
                        reader2 = cmd1e.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2.Read

                            Dim sid = reader2.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2.Close()
                        With officeexcel
                            .Range("G14".ToString).Value = s1
                        End With

                        Dim cmd1ea As New SqlCommand("select [" + Format(DateTimePicker3.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2a As SqlDataReader
                        reader2a = cmd1ea.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2a.Read

                            Dim sid = reader2a.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2a.Close()
                        With officeexcel
                            .Range("G15".ToString).Value = s1
                        End With

                        Dim cmd1eb As New SqlCommand("select [" + Format(DateTimePicker4.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2b As SqlDataReader
                        reader2b = cmd1eb.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2b.Read

                            Dim sid = reader2b.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2b.Close()
                        With officeexcel
                            .Range("G16".ToString).Value = s1
                        End With

                        Dim cmd1ec As New SqlCommand("select [" + Format(DateTimePicker5.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2c As SqlDataReader
                        reader2c = cmd1ec.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2c.Read

                            Dim sid = reader2c.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2c.Close()
                        With officeexcel
                            .Range("G17".ToString).Value = s1
                        End With

                        Dim cmd1ed As New SqlCommand("select [" + Format(DateTimePicker6.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2d As SqlDataReader
                        reader2d = cmd1ed.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2d.Read

                            Dim sid = reader2d.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2d.Close()
                        With officeexcel
                            .Range("G18".ToString).Value = s1
                        End With

                        Dim cmd1ee As New SqlCommand("select [" + Format(DateTimePicker7.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2e As SqlDataReader
                        reader2e = cmd1ee.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2e.Read

                            Dim sid = reader2e.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2e.Close()
                        With officeexcel
                            .Range("G19".ToString).Value = s1
                        End With

                        Dim cmd1ef As New SqlCommand("select [" + Format(DateTimePicker8.Value, "dd-MM-yyyy") + "] from otltg", con7)
                        Dim reader2f As SqlDataReader
                        reader2f = cmd1ef.ExecuteReader

                        s1 = 0
                        b1 = 0

                        While reader2f.Read

                            Dim sid = reader2f.GetValue(0)
                            If sid >= 0 Then
                                s1 = s1 + sid
                            Else
                                b1 = b1 + sid
                            End If

                        End While
                        reader2f.Close()
                        With officeexcel
                            .Range("G20".ToString).Value = s1
                        End With

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If

            Try

                For i As Integer = 0 To ds.Tables("salary").Rows.Count - 1
                    With officeexcel

                        .Range("B" + (i + 25).ToString).Value = ds.Tables("salary").Rows(i).Item(1).ToString
                        .Range("C" + (i + 25).ToString).Value = ds.Tables("salary").Rows(i).Item(2).ToString
                        .Range("D" + (i + 25).ToString).Value = ds.Tables("salary").Rows(i).Item(3).ToString

                        .Range("F" + (i + 25).ToString).Value = ds.Tables("salary").Rows(i).Item(5).ToString

                        Dim cmdz As New SqlCommand("select salary from register where id='" + ds.Tables("salary").Rows(i).Item(0).ToString + "'", con7)
                        Dim dr As SqlDataReader
                        dr = cmdz.ExecuteReader
                        dr.Read()
                        TextBox1.Text = dr.GetString(0)
                        dr.Close()

                        TextBox2.Text = ds.Tables("salary").Rows(i).Item(7).ToString / 8

                        s1 = Val(TextBox1.Text) * TextBox2.Text

                        .Range("H" + (i + 25).ToString).Value = s1

                        j1 = i
                    End With
                Next

                With officeexcel

                    .Range("B" + (j1 + 27).ToString).Value = "Total"

                End With

                With officeexcel

                    .Range("A" + (j1 + 32).ToString).Value = "Maintenance Manager"
                    .Range("C" + (j1 + 32).ToString).Value = "Administrative Officer"
                    .Range("F" + (j1 + 32).ToString).Value = "Principal"
                    .Range("H" + (j1 + 32).ToString).Value = "Secretary"
                End With

                officeexcel = Nothing
                workbook = Nothing

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            MessageBox.Show("Fetch some data to export")
        End If

    End Sub
    Private Sub Form7_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Form1.Show()
    End Sub
End Class