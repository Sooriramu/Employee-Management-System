Imports System.Data
Imports System.Data.SqlClient
Public Class Form5
    Dim con4 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim command As New SqlCommand

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If cb1.Text = "Men mazdoor" Then
                If cb2.Text <> "" Then

                    Dim result As Integer = MessageBox.Show("Do you want to delete the " + Label4.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

                    If result = DialogResult.OK Then

                        Dim cmd3f As New SqlCommand("insert into history values('" + cb2.Text + "','" + cb1.Text + "' ,'" + Label4.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "' )  ", con4)
                        cmd3f.ExecuteNonQuery()

                        Dim cmd3a As New SqlCommand("delete from amm where id='" + cb2.Text + "'  ", con4)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("delete from otltmm where id='" + cb2.Text + "'  ", con4)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Deleted")
                        Call Button4_Click(sender, e)

                    End If
                Else
                    MessageBox.Show("Select the ID")
                End If

            ElseIf cb1.Text = "Carpenter" Then
                If cb2.Text <> "" Then


                    Dim result As Integer = MessageBox.Show("Do you want to delete the " + Label4.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

                    If result = DialogResult.OK Then

                        Dim cmd3f As New SqlCommand("insert into history values('" + cb2.Text + "','" + cb1.Text + "' ,'" + Label4.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "' )  ", con4)
                        cmd3f.ExecuteNonQuery()

                        Dim cmd3a As New SqlCommand("delete from ac where id='" + cb2.Text + "'  ", con4)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("delete from otltc where id='" + cb2.Text + "'  ", con4)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Deleted")
                        Call Button4_Click(sender, e)
                    End If
                Else
                    MessageBox.Show("Select the ID")
                End If

            ElseIf cb1.Text = "Women mazdoor" Then
                If cb2.Text <> "" Then


                    Dim result As Integer = MessageBox.Show("Do you want to delete the " + Label4.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

                    If result = DialogResult.OK Then

                        Dim cmd3f As New SqlCommand("insert into history values('" + cb2.Text + "','" + cb1.Text + "' ,'" + Label4.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "' )  ", con4)
                        cmd3f.ExecuteNonQuery()

                        Dim cmd3a As New SqlCommand("delete from awm where id='" + cb2.Text + "'  ", con4)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("delete from otltwm where id='" + cb2.Text + "'  ", con4)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Deleted")
                        Call Button4_Click(sender, e)
                    End If
                Else
                    MessageBox.Show("Select the ID")

                End If

            ElseIf cb1.Text = "Mason" Then
                If cb2.Text <> "" Then


                    Dim result As Integer = MessageBox.Show("Do you want to delete the " + Label4.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

                    If result = DialogResult.OK Then

                        Dim cmd3f As New SqlCommand("insert into history values('" + cb2.Text + "','" + cb1.Text + "' ,'" + Label4.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "' )  ", con4)
                        cmd3f.ExecuteNonQuery()

                        Dim cmd3a As New SqlCommand("delete from am where id='" + cb2.Text + "'  ", con4)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("delete from otltm where id='" + cb2.Text + "'  ", con4)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Deleted")
                        Call Button4_Click(sender, e)
                    End If
                Else
                    MessageBox.Show("Select the ID")
                End If

            ElseIf cb1.Text = "Gardener" Then
                If cb2.Text <> "" Then


                    Dim result As Integer = MessageBox.Show("Do you want to delete the " + Label4.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

                    If result = DialogResult.OK Then

                        Dim cmd3f As New SqlCommand("insert into history values('" + cb2.Text + "','" + cb1.Text + "' ,'" + Label4.Text + "','" + Format(DateTime.Now, "yyyy-MM-dd") + "' )  ", con4)
                        cmd3f.ExecuteNonQuery()

                        Dim cmd3a As New SqlCommand("delete from ag where id='" + cb2.Text + "'  ", con4)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("delete from otltg where id='" + cb2.Text + "'  ", con4)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Deleted")
                        Call Button4_Click(sender, e)
                    End If
                Else
                    MessageBox.Show("Select the ID")
                End If

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        cb2.Items.Clear()
        Label4.Text = ""
        Dim cmd1b As New SqlDataAdapter
        Dim dt As New DataTable
        Dim query As String
        Dim bsource As New BindingSource
        Try
            If cb1.Text = "Men mazdoor" Then
                query = "select id,name from amm order by name asc"
            ElseIf cb1.Text = "Carpenter" Then
                query = "select id,name from ac order by name asc"
            ElseIf cb1.Text = "Women mazdoor" Then
                query = "select id,name from awm order by name asc"
            ElseIf cb1.Text = "Mason" Then
                query = "select id,name from am order by name asc"
            ElseIf cb1.Text = "Gardener" Then
                query = "select id,name from ag order by name asc"
            End If
            command = New SqlCommand(query, con4)
            cmd1b.SelectCommand = command
            cmd1b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1b.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from amm", con4)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader1.Close()

            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1d As New SqlCommand("select id from ac", con4)
                Dim reader2 As SqlDataReader
                reader2 = cmd1d.ExecuteReader

                While reader2.Read
                    Dim sid = reader2.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader2.Close()

            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from awm", con4)
                Dim reader3 As SqlDataReader
                reader3 = cmd1d.ExecuteReader

                While reader3.Read
                    Dim sid = reader3.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader3.Close()

            ElseIf cb1.Text = "Mason" Then
                Dim cmd1d As New SqlCommand("select id from am", con4)
                Dim reader4 As SqlDataReader
                reader4 = cmd1d.ExecuteReader

                While reader4.Read
                    Dim sid = reader4.GetString(0)
                    cb2.Items.Add(sid)
                End While

                reader4.Close()

            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1d As New SqlCommand("select id from ag", con4)
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

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        Call Button4_Click(sender, e)

    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        con4.Open()

        cb1.SelectedIndex = 0

        Label4.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()


    End Sub

    Private Sub cb2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb2.SelectedIndexChanged
        Try
            If cb1.Text = "Men mazdoor" Then
                Dim cmd1y As New SqlCommand("select Name from amm where id='" + cb2.Text + "'", con4)
                Dim dr As SqlDataReader
                dr = cmd1y.ExecuteReader
                dr.Read()
                Label4.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1z As New SqlCommand("select Name from ac where id='" + cb2.Text + "'", con4)
                Dim dr As SqlDataReader
                dr = cmd1z.ExecuteReader
                dr.Read()
                Label4.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1aa As New SqlCommand("select Name from awm where id='" + cb2.Text + "'", con4)
                Dim dr As SqlDataReader
                dr = cmd1aa.ExecuteReader
                dr.Read()
                Label4.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1ab As New SqlCommand("select Name from am where id='" + cb2.Text + "'", con4)
                Dim dr As SqlDataReader
                dr = cmd1ab.ExecuteReader
                dr.Read()
                Label4.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1ac As New SqlCommand("select Name from ag where id='" + cb2.Text + "'", con4)
                Dim dr As SqlDataReader
                dr = cmd1ac.ExecuteReader
                dr.Read()
                Label4.Text = dr.GetValue(0)
                dr.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Form5_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Hide()
        Form1.Show()
    End Sub

End Class