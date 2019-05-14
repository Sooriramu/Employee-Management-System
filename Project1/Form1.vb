Imports System.Data
Imports System.Data.SqlClient


Public Class Form1
    Public Property str As String
    Dim a = 0
    Dim b = 0
    Dim c = 0
    Dim d = 0
    Dim e = 0
    Dim date1
    Dim n6, n7
    Dim command As SqlCommand

    Dim con As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.Open()
        DateTimePicker1.Enabled = False
        cb2.Items.Clear()
        Button5.Hide()
        Button6.Hide()
        Label4.Text = str
        DateTimePicker1.Value = DateTime.Now
        cb1.SelectedIndex = 0
        cb2.SelectedIndex = -1
        cb3.SelectedIndex = 0
        cb4.SelectedIndex = 0
        cb5.SelectedIndex = 0
        Label12.Text = ""

        date1 = Date.Now.ToShortDateString
        Try
            Dim cmd1c As New SqlCommand("alter table amm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1c.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1f As New SqlCommand("alter table otmm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1f.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1g As New SqlCommand("alter table ltmm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1g.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1l As New SqlCommand("alter table ac add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1l.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1m As New SqlCommand("alter table otc add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1m.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1n As New SqlCommand("alter table ltc add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1n.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1o As New SqlCommand("alter table awm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1o.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1p As New SqlCommand("alter table otwm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1p.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1q As New SqlCommand("alter table ltwm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1q.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1r As New SqlCommand("alter table am add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1r.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1s As New SqlCommand("alter table otm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1s.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1t As New SqlCommand("alter table ltm add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1t.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1u As New SqlCommand("alter table ag add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1u.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1v As New SqlCommand("alter table otg add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1v.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Try
            Dim cmd1w As New SqlCommand("alter table ltg add [" + DateTime.Now.ToShortDateString + "] varchar(15) default(0)", con)
            cmd1w.ExecuteNonQuery()
        Catch ex As Exception
        End Try

        Call refresh_click(sender, e)

    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged

        cb2.Items.Clear()
        Label12.Text = ""
        If cb1.Text = "Men mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from amm", con)
            Dim reader1 As SqlDataReader
            reader1 = cmd1d.ExecuteReader

            While reader1.Read
                Dim sid = reader1.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader1.Close()
            Call refresh_click(sender, e)
        ElseIf cb1.text = "Carpenter" Then
            Dim cmd1d As New SqlCommand("select id from ac", con)
            Dim reader2 As SqlDataReader
            reader2 = cmd1d.ExecuteReader

            While reader2.Read
                Dim sid = reader2.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader2.Close()
            Call refresh_click(sender, e)
        ElseIf cb1.text = "Women mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from awm", con)
            Dim reader3 As SqlDataReader
            reader3 = cmd1d.ExecuteReader

            While reader3.Read
                Dim sid = reader3.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader3.Close()
            Call refresh_click(sender, e)
        ElseIf cb1.text = "Mason" Then
            Dim cmd1d As New SqlCommand("select id from am", con)
            Dim reader4 As SqlDataReader
            reader4 = cmd1d.ExecuteReader

            While reader4.Read
                Dim sid = reader4.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader4.Close()
            Call refresh_click(sender, e)
        ElseIf cb1.text = "Gardener" Then
            Dim cmd1d As New SqlCommand("select id from ag", con)
            Dim reader5 As SqlDataReader
            reader5 = cmd1d.ExecuteReader

            While reader5.Read
                Dim sid = reader5.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader5.Close()
            Call refresh_click(sender, e)
        End If
    End Sub

    Private Sub view_Click(sender As Object, e As EventArgs) Handles view.Click

        Dim cmd1b As New SqlDataAdapter
            Dim dt As New DataTable
            Dim query As String
            Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select * from amm"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select * from ac"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select * from awm"
            ElseIf cb1.Text = "Mason" Then
                query = "select * from am"
            ElseIf cb1.Text = "Gardener" Then
                query = "select * from ag"
            End If
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If cb2.Text <> "" Then
            If cb1.Text = "Men mazdoor" Then
                Dim cmd1e As New SqlCommand("update amm set [" + DateTime.Now.ToShortDateString + "]='" + cb3.Text + "' where ID='" + cb2.Text + "'", con)
                cmd1e.ExecuteNonQuery()

                Dim cmd1h As New SqlCommand("update otmm set [" + DateTime.Now.ToShortDateString + "]='" + cb4.Text + "' where ID='" + cb2.Text + "' ", con)
                cmd1h.ExecuteNonQuery()

                Dim cmd1i As New SqlCommand("update ltmm set [" + DateTime.Now.ToShortDateString + "]='" + cb5.Text + "' where ID='" + cb2.Text + "'", con)
                cmd1i.ExecuteNonQuery()

            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1e As New SqlCommand("update ac set [" + DateTime.Now.ToShortDateString + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                cmd1e.ExecuteNonQuery()

                Dim cmd1h As New SqlCommand("update otc set [" + DateTime.Now.ToShortDateString + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                cmd1h.ExecuteNonQuery()

                Dim cmd1i As New SqlCommand("update ltc set [" + DateTime.Now.ToShortDateString + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                cmd1i.ExecuteNonQuery()

            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1e As New SqlCommand("update awm set [" + DateTime.Now.ToShortDateString + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                cmd1e.ExecuteNonQuery()

                Dim cmd1h As New SqlCommand("update otwm set [" + DateTime.Now.ToShortDateString + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                cmd1h.ExecuteNonQuery()

                Dim cmd1i As New SqlCommand("update ltwm set [" + DateTime.Now.ToShortDateString + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                cmd1i.ExecuteNonQuery()

            ElseIf cb1.Text = "Mason" Then
                Dim cmd1e As New SqlCommand("update am set [" + DateTime.Now.ToShortDateString + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                cmd1e.ExecuteNonQuery()

                Dim cmd1h As New SqlCommand("update otm set [" + DateTime.Now.ToShortDateString + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                cmd1h.ExecuteNonQuery()

                Dim cmd1i As New SqlCommand("update ltm set [" + DateTime.Now.ToShortDateString + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                cmd1i.ExecuteNonQuery()

            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1e As New SqlCommand("update ag set [" + DateTime.Now.ToShortDateString + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                cmd1e.ExecuteNonQuery()

                Dim cmd1h As New SqlCommand("update otg set [" + DateTime.Now.ToShortDateString + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                cmd1h.ExecuteNonQuery()

                Dim cmd1i As New SqlCommand("update ltg set [" + DateTime.Now.ToShortDateString + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                cmd1i.ExecuteNonQuery()

            End If
            MessageBox.Show("Updated")
            Call refresh_click(sender, e)
        Else
            MessageBox.Show("Select the ID")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cmd1j As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select * from otmm"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select * from otc"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select * from otwm"
            ElseIf cb1.Text = "Mason" Then
                query = "select * from otm"
            ElseIf cb1.Text = "Gardener" Then
                query = "select * from otg"
            End If
            command = New SqlCommand(query, con)
            cmd1j.SelectCommand = command
            cmd1j.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1j.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim cmd1k As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select * from ltmm"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select * from ltc"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select * from ltwm"
            ElseIf cb1.Text = "Mason" Then
                query = "select * from ltm"
            ElseIf cb1.Text = "Gardener" Then
                query = "select * from ltg"
            End If
            command = New SqlCommand(query, con)
            cmd1k.SelectCommand = command
            cmd1k.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1k.Update(dt)

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

                    Dim cmd1h As New SqlCommand("update otmm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                    Dim cmd1i As New SqlCommand("update ltmm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                    cmd1i.ExecuteNonQuery()

                ElseIf cb1.Text = "Carpenter" Then
                    Dim cmd1e As New SqlCommand("update ac set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "'", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otc set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                    Dim cmd1i As New SqlCommand("update ltc set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                    cmd1i.ExecuteNonQuery()

                ElseIf cb1.Text = "Women mazdoor" Then
                    Dim cmd1e As New SqlCommand("update awm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otwm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                    Dim cmd1i As New SqlCommand("update ltwm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                    cmd1i.ExecuteNonQuery()

                ElseIf cb1.Text = "Mason" Then
                    Dim cmd1e As New SqlCommand("update am set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                    Dim cmd1i As New SqlCommand("update ltm set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                    cmd1i.ExecuteNonQuery()

                ElseIf cb1.Text = "Gardener" Then
                    Dim cmd1e As New SqlCommand("update ag set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb3.Text + "' where ID='" + cb2.Text + "' ", con)
                    cmd1e.ExecuteNonQuery()

                    Dim cmd1h As New SqlCommand("update otg set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb4.Text + "' where  ID='" + cb2.Text + "' ", con)
                    cmd1h.ExecuteNonQuery()

                    Dim cmd1i As New SqlCommand("update ltg set [" + Format(DateTimePicker1.Value, "dd-MM-yyyy") + "]='" + cb5.Text + "' where  ID='" + cb2.Text + "'", con)
                    cmd1i.ExecuteNonQuery()

                End If
                MessageBox.Show("Updated")
                Button6.Hide()
                Button5.Hide()
                Button1.Show()
                DateTimePicker1.Enabled = False
                Call refresh_click(sender, e)
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
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub refresh_Click(sender As Object, e As EventArgs) Handles refresh.Click
        cb2.Items.Clear()
        cb3.SelectedIndex = 0
        cb4.SelectedIndex = 0
        cb5.SelectedIndex = 0
        Label12.Text = ""
        DateTimePicker1.Value = DateTime.Now

        If cb1.Text = "Men mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from amm", con)
            Dim reader1 As SqlDataReader
            reader1 = cmd1d.ExecuteReader

            While reader1.Read
                Dim sid = reader1.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader1.Close()
        ElseIf cb1.Text = "Carpenter" Then
            Dim cmd1d As New SqlCommand("select id from ac", con)
            Dim reader2 As SqlDataReader
            reader2 = cmd1d.ExecuteReader

            While reader2.Read
                Dim sid = reader2.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader2.Close()
        ElseIf cb1.Text = "Women mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from awm", con)
            Dim reader3 As SqlDataReader
            reader3 = cmd1d.ExecuteReader

            While reader3.Read
                Dim sid = reader3.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader3.Close()
        ElseIf cb1.Text = "Mason" Then
            Dim cmd1d As New SqlCommand("select id from am", con)
            Dim reader4 As SqlDataReader
            reader4 = cmd1d.ExecuteReader

            While reader4.Read
                Dim sid = reader4.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader4.Close()
        ElseIf cb1.Text = "Gardener" Then
            Dim cmd1d As New SqlCommand("select id from ag", con)
            Dim reader5 As SqlDataReader
            reader5 = cmd1d.ExecuteReader

            While reader5.Read
                Dim sid = reader5.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader5.Close()
        End If

        Dim cmd1b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select * from amm"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select * from ac"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select * from awm"
            ElseIf cb1.Text = "Mason" Then
                query = "select * from am"
            ElseIf cb1.Text = "Gardener" Then
                query = "select * from ag"
            End If
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

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

        Call refresh_Click(sender, e)
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
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Form5.Show()
    End Sub

    Private Sub cb2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb2.SelectedIndexChanged
        cb3.SelectedIndex = 0
        cb4.SelectedIndex = 0
        cb5.SelectedIndex = 0
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

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        con.Close()
    End Sub
End Class
