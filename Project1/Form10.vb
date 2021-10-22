Imports System.Data.SqlClient

Public Class Form10
    Dim command As SqlCommand
    Dim con10 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con10.Open()
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        cb1.Items.Clear()
        Try
            Dim cmd1j As New SqlDataAdapter
            Dim dt As New DataTable
            Dim query1 As String
            Dim bsource As New BindingSource

            query1 = "select username from login where username <>'psg'"

            command = New SqlCommand(query1, con10)
            cmd1j.SelectCommand = command
            cmd1j.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1j.Update(dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try

        Try

            Dim cmd1d As New SqlCommand("select username from login where username <>'psg'", con10)
            Dim reader1 As SqlDataReader
            reader1 = cmd1d.ExecuteReader

            While reader1.Read
                Dim sid = reader1.GetString(0)
                cb1.Items.Add(sid)
            End While
            reader1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
        cb1.SelectedIndex = 0

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim result As Integer = MessageBox.Show("Do you want to delete the " + cb1.Text + "'s record?", "Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

        If result = DialogResult.OK Then

            Dim cmd3a As New SqlCommand("delete from login where username='" + cb1.Text + "'  ", con10)
            cmd3a.ExecuteNonQuery()

            MessageBox.Show("Deleted successfully")

            cb1.Items.Clear()

            Try
                Dim cmd1j As New SqlDataAdapter
                Dim dt As New DataTable
                Dim query1 As String
                Dim bsource As New BindingSource

                query1 = "select username from login where username <>'psg'"

                command = New SqlCommand(query1, con10)
                cmd1j.SelectCommand = command
                cmd1j.Fill(dt)
                bsource.DataSource = dt
                DataGridView1.DataSource = bsource
                cmd1j.Update(dt)
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString)
            End Try

            Try

                Dim cmd1d As New SqlCommand("select username from login where username <>'psg'", con10)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb1.Items.Add(sid)
                End While
                reader1.Close()

            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString)
            End Try
            cb1.SelectedIndex = 0

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form10_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Hide()
        Form1.Show()
    End Sub
End Class