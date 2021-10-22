Imports System.Data
Imports System.Data.SqlClient

Public Class Form1

    Dim a = 0
    Dim b = 0
    Dim c = 0
    Dim d = 0
    Dim e = 0
    Dim i = 0
    Dim n6, n7,n2,n3
    Dim n8 = 1
    Dim command As SqlCommand
    Dim con As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.Open()

        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)

        DateTimePicker1.Enabled = False
        cb2.Items.Clear()
        Button5.Hide()
        Button6.Hide()
        Label13.Text = "ATTENDANCE"

        DateTimePicker1.Value = DateTime.Now
        cb1.SelectedIndex = 0
        cb2.SelectedIndex = -1
        cb3.SelectedIndex = 0
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        Label12.Text = ""

        Try
            Dim cmd1c As New SqlCommand("alter table amm add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0 ", con)
            cmd1c.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1f As New SqlCommand("alter table otltmm add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1f.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1l As New SqlCommand("alter table ac add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1l.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1m As New SqlCommand("alter table otltc add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1m.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1o As New SqlCommand("alter table awm add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1o.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1p As New SqlCommand("alter table otltwm add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1p.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1r As New SqlCommand("alter table am add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1r.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1s As New SqlCommand("alter table otltm add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1s.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1u As New SqlCommand("alter table ag add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1u.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1v As New SqlCommand("alter table otltg add [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "] float Not null default 0", con)
            cmd1v.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Call refresh_Click(sender, e)

    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged

        cb2.Items.Clear()
        Label12.Text = ""
        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from amm", con)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader1.Close()
                Call refresh_Click(sender, e)
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1d As New SqlCommand("select id from ac", con)
                Dim reader2 As SqlDataReader
                reader2 = cmd1d.ExecuteReader

                While reader2.Read
                    Dim sid = reader2.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader2.Close()
                Call refresh_Click(sender, e)
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from awm", con)
                Dim reader3 As SqlDataReader
                reader3 = cmd1d.ExecuteReader

                While reader3.Read
                    Dim sid = reader3.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader3.Close()
                Call refresh_Click(sender, e)
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1d As New SqlCommand("select id from am", con)
                Dim reader4 As SqlDataReader
                reader4 = cmd1d.ExecuteReader

                While reader4.Read
                    Dim sid = reader4.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader4.Close()
                Call refresh_Click(sender, e)
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1d As New SqlCommand("select id from ag", con)
                Dim reader5 As SqlDataReader
                reader5 = cmd1d.ExecuteReader

                While reader5.Read
                    Dim sid = reader5.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader5.Close()
                Call refresh_Click(sender, e)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If cb2.Text <> "" Then
            Try
                If cb1.Text = "Men mazdoor" Then
                    Dim cmd1e As New SqlCommand("update amm set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "'", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltmm set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + tb1.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Carpenter" Then
                    Dim cmd1e As New SqlCommand("update ac set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltc set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Women mazdoor" Then
                    Dim cmd1e As New SqlCommand("update awm set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltwm set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Mason" Then
                    Dim cmd1e As New SqlCommand("update am set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltm set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Gardener" Then
                    Dim cmd1e As New SqlCommand("update ag set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltg set [" + DateTime.Now.Date.ToString("dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            MessageBox.Show("Updated")
            Call refresh_Click(sender, e)

        Else
            MessageBox.Show("Select the ID")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Label13.Text = "OT/LT"
        Dim cmd1j As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query1 As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query1 = "select id,name, [" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "] from otltmm order by len(id), id"
            ElseIf cb1.Text = "Carpenter" Then
                query1 = "select id,name, [" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "] from otltc order by len(id), id"
            ElseIf cb1.Text = "Women mazdoor" Then
                query1 = "select id,name, [" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "] from otltwm order by len(id), id"
            ElseIf cb1.Text = "Mason" Then
                query1 = "select id,name, [" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "] from otltm order by len(id), id"
            ElseIf cb1.Text = "Gardener" Then
                query1 = "select id,name, [" + DateTimePicker1.Value.Date.ToString("dd-MM-yyyy") + "] from otltg order by len(id), id"
            End If
            command = New SqlCommand(query1, con)
            cmd1j.SelectCommand = command
            cmd1j.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1j.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If cb2.Text <> "" Then
            Try
                If cb1.Text = "Men mazdoor" Then
                    Dim cmd1e As New SqlCommand("update amm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltmm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Carpenter" Then
                    Dim cmd1e As New SqlCommand("update ac set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "'", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltc set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Women mazdoor" Then
                    Dim cmd1e As New SqlCommand("update awm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltwm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Mason" Then
                    Dim cmd1e As New SqlCommand("update am set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                ElseIf cb1.Text = "Gardener" Then
                    Dim cmd1e As New SqlCommand("update ag set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otltg set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + tb1.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                End If
                MessageBox.Show("Updated")

                Call refresh_Click(sender, e)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            MessageBox.Show("Select the ID")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button6.Hide()
        Button5.Hide()
        Button1.Show()
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker1.Enabled = False
        Call refresh_Click(sender, e)
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub refresh_Click(sender As Object, e As EventArgs) Handles refresh.Click

        cb2.Items.Clear()
        cb3.SelectedIndex = 0
        tb1.Text = 0
        Label12.Text = ""
        Label13.Text = "ATTENDANCE"
        DateTimePicker1.Value = DateTime.Now

        Dim cmd1b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query1 As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query1 = "select id,name, [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm order by len(id), id"
            ElseIf cb1.Text = "Carpenter" Then
                query1 = "select id,name, [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac order by len(id), id"
            ElseIf cb1.Text = "Women mazdoor" Then
                query1 = "select id,name, [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm order by len(id), id"
            ElseIf cb1.Text = "Mason" Then
                query1 = "select id,name, [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am order by len(id), id"
            ElseIf cb1.Text = "Gardener" Then
                query1 = "select id,name, [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag order by len(id), id"
            End If
            command = New SqlCommand(query1, con)
            cmd1b.SelectCommand = command
            cmd1b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1b.Update(dt)

        Catch ex As Exception

        End Try

        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from amm ", con)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader1.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1d As New SqlCommand("select id from ac", con)
                Dim reader2 As SqlDataReader
                reader2 = cmd1d.ExecuteReader

                While reader2.Read
                    Dim sid = reader2.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader2.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from awm", con)
                Dim reader3 As SqlDataReader
                reader3 = cmd1d.ExecuteReader

                While reader3.Read
                    Dim sid = reader3.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader3.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1d As New SqlCommand("select id from am", con)
                Dim reader4 As SqlDataReader
                reader4 = cmd1d.ExecuteReader

                While reader4.Read
                    Dim sid = reader4.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader4.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1d As New SqlCommand("select id from ag", con)
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

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

        Call refresh_Click(sender, e)
        Try
            Dim cmd1x As New SqlCommand("select pin from login where username='psg'", con)
            Dim dr As SqlDataReader
            dr = cmd1x.ExecuteReader
            dr.Read()
            n6 = dr.GetValue(0)
            dr.Close()
            n7 = InputBox("Enter pin", "Security check")
            If n6 = n7 Then
                Button5.Show()
                Button6.Show()
                Button1.Hide()
                DateTimePicker1.Enabled = True
            Else
                MessageBox.Show("Invalid pin")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        Form5.Show()

    End Sub

    Private Sub cb2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb2.SelectedIndexChanged
        cb3.SelectedIndex = 0


        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1y As New SqlCommand("select Name from amm where id='" + cb2.Text + "'", con)
                Dim dr As SqlDataReader
                dr = cmd1y.ExecuteReader
                dr.Read()
                Label12.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1z As New SqlCommand("select Name from ac where id='" + cb2.Text + "'", con)
                Dim dr As SqlDataReader
                dr = cmd1z.ExecuteReader
                dr.Read()
                Label12.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1aa As New SqlCommand("select Name from awm where id='" + cb2.Text + "'", con)
                Dim dr As SqlDataReader
                dr = cmd1aa.ExecuteReader
                dr.Read()
                Label12.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1ab As New SqlCommand("select Name from am where id='" + cb2.Text + "'", con)
                Dim dr As SqlDataReader
                dr = cmd1ab.ExecuteReader
                dr.Read()
                Label12.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1ac As New SqlCommand("select Name from ag where id='" + cb2.Text + "'", con)
                Dim dr As SqlDataReader
                dr = cmd1ac.ExecuteReader
                dr.Read()
                Label12.Text = dr.GetValue(0)
                dr.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        Try
            If cb1.Text = "Men mazdoor" Then
                Dim cmd1y As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltmm where id='" + cb2.Text + "'", con)
                Dim dr1 As SqlDataReader
                dr1 = cmd1y.ExecuteReader
                dr1.Read()
                tb1.Text = dr1.GetValue(0)
                dr1.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1z As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltc where id='" + cb2.Text + "'", con)
                Dim dr1 As SqlDataReader
                dr1 = cmd1z.ExecuteReader
                dr1.Read()
                tb1.Text = dr1.GetValue(0)
                dr1.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltwm where id='" + cb2.Text + "'", con)
                Dim dr1 As SqlDataReader
                dr1 = cmd1aa.ExecuteReader
                dr1.Read()
                tb1.Text = dr1.GetValue(0)
                dr1.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1ab As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltm where id='" + cb2.Text + "'", con)
                Dim dr1 As SqlDataReader
                dr1 = cmd1ab.ExecuteReader
                dr1.Read()
                tb1.Text = dr1.GetValue(0)
                dr1.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1ac As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from otltg where id='" + cb2.Text + "'", con)
                Dim dr1 As SqlDataReader
                dr1 = cmd1ac.ExecuteReader
                dr1.Read()
                tb1.Text = dr1.GetValue(0)
                dr1.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1y As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from amm where id='" + cb2.Text + "'", con)
                Dim dr2 As SqlDataReader
                dr2 = cmd1y.ExecuteReader
                dr2.Read()
                cb3.Text = dr2.GetValue(0)
                dr2.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1z As New SqlCommand("select[" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ac where id='" + cb2.Text + "'", con)
                Dim dr2 As SqlDataReader
                dr2 = cmd1z.ExecuteReader
                dr2.Read()
                cb3.Text = dr2.GetValue(0)
                dr2.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1aa As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from awm where id='" + cb2.Text + "'", con)
                Dim dr2 As SqlDataReader
                dr2 = cmd1aa.ExecuteReader
                dr2.Read()
                cb3.Text = dr2.GetValue(0)
                dr2.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1ab As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from am where id='" + cb2.Text + "'", con)
                Dim dr2 As SqlDataReader
                dr2 = cmd1ab.ExecuteReader
                dr2.Read()
                cb3.Text = dr2.GetValue(0)
                dr2.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1ac As New SqlCommand("select [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "] from ag where id='" + cb2.Text + "'", con)
                Dim dr2 As SqlDataReader
                dr2 = cmd1ac.ExecuteReader
                dr2.Read()
                cb3.Text = dr2.GetValue(0)
                dr2.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Label13.Text = "HISTORY"

        Dim cmd1b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            query = "select * from history order by date desc "
            command = New SqlCommand(query, con)
            cmd1b.SelectCommand = command
            cmd1b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1b.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        con.Close()
        Application.Exit()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Hide()
        Form6.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Label13.Text = "DB"

        Dim cmd1b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            query = "select * from register order by DOJ desc"
            command = New SqlCommand(query, con)
            cmd1b.SelectCommand = command
            cmd1b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1b.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ContactToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactToolStripMenuItem.Click
        Me.Hide()
        Form9.Show()
    End Sub



    Private Sub DailyRecordsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DailyRecordsToolStripMenuItem1.Click
        Me.Hide()
        Form8.Show()
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub cb3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb3.SelectedIndexChanged

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub tb1_TextChanged(sender As Object, e As EventArgs) Handles tb1.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub SalaryPrintoutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SalaryPrintoutToolStripMenuItem1.Click
        Me.Hide()
        Form7.Show()
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

        Dim cmd2b As New SqlCommand("select pin from login where username='psg'", con)
        Dim dr As SqlDataReader
        dr = cmd2b.ExecuteReader
        dr.Read()
        n2 = dr.GetValue(0)
        dr.Close()
        n3 = InputBox("Enter pin", "Security check")
        If n2 = n3 Then
            Me.Hide()
            Form10.Show()
        Else
            MessageBox.Show("Invalid pin")
        End If

    End Sub

    Private Sub tb1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb1.KeyPress
        Try
            Dim dac As String = ",<.>/?;:'[{]}=+_()*&^%$#@!~`\| "

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
