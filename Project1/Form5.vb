Imports System.Data
Imports System.Data.SqlClient
Public Class Form5
    Dim con4 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim command As New SqlCommand

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If cb1.Text = "Men mazdoor" Then
            If cb2.Text <> "" Then

                Dim result As Integer = MessageBox.Show("Do you want to delete the record?", "Alert", MessageBoxButtons.OKCancel)

                If result = DialogResult.OK Then

                    Dim cmd3a As New SqlCommand("delete from amm where id='" + cb2.Text + "'  ", con4)
                    cmd3a.ExecuteNonQuery()

                    Dim cmd3d As New SqlCommand("delete from otmm where id='" + cb2.Text + "'  ", con4)
                    cmd3d.ExecuteNonQuery()

                    Dim cmd3e As New SqlCommand("delete from ltmm where id='" + cb2.Text + "'  ", con4)
                    cmd3e.ExecuteNonQuery()

                    MessageBox.Show("Deleted")
                    Call Button4_Click(sender, e)

                End If
                Else
                MessageBox.Show("Select the ID")
            End If

        ElseIf cb1.Text = "Carpenter" Then
            If cb2.Text <> "" Then

                Dim cmd3a As New SqlCommand("delete from ac where id='" + cb2.Text + "'  ", con4)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("delete from otc where id='" + cb2.Text + "'  ", con4)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("delete from ltc where id='" + cb2.Text + "'  ", con4)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Deleted")
                Call Button4_Click(sender, e)
            Else
                MessageBox.Show("Select the ID")
            End If

        ElseIf cb1.Text = "Women mazdoor" Then
            If cb2.Text <> "" Then

                Dim cmd3a As New SqlCommand("delete from awm where id='" + cb2.Text + "'  ", con4)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("delete from otwm where id='" + cb2.Text + "'  ", con4)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("delete from ltwm where id='" + cb2.Text + "'  ", con4)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Deleted")
                Call Button4_Click(sender, e)
            Else
                MessageBox.Show("Select the ID")
            End If

        ElseIf cb1.Text = "Mason" Then
            If cb2.Text <> "" Then

                Dim cmd3a As New SqlCommand("delete from am where id='" + cb2.Text + "'  ", con4)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("delete from otm where id='" + cb2.Text + "'  ", con4)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("delete from ltm where id='" + cb2.Text + "'  ", con4)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Deleted")
                Call Button4_Click(sender, e)
            Else
                MessageBox.Show("Select the ID")
            End If

        ElseIf cb1.Text = "Gardener" Then
            If cb2.Text <> "" Then

                Dim cmd3a As New SqlCommand("delete from ag where id='" + cb2.Text + "'  ", con4)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("delete from otg where id='" + cb2.Text + "'  ", con4)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("delete from ltg where id='" + cb2.Text + "'  ", con4)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Deleted")
                Call Button4_Click(sender, e)
            Else
                MessageBox.Show("Select the ID")
            End If

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        cb2.Items.Clear()
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
            command = New SqlCommand(query, con4)
            cmd1b.SelectCommand = command
            cmd1b.Fill(dt)
            bsource.DataSource = dt
            DataGridView1.DataSource = bsource
            cmd1b.Update(dt)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        If cb1.Text = "Men mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from amm", con4)
            Dim reader1 As SqlDataReader
            reader1 = cmd1d.ExecuteReader

            While reader1.Read
                Dim sid = reader1.GetInt32(0)
                cb2.Items.Add(sid)
            End While
            reader1.Close()

        ElseIf cb1.Text = "Carpenter" Then
            Dim cmd1d As New SqlCommand("select id from ac", con4)
            Dim reader2 As SqlDataReader
            reader2 = cmd1d.ExecuteReader

            While reader2.Read
                Dim sid = reader2.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader2.Close()

        ElseIf cb1.Text = "Women mazdoor" Then
            Dim cmd1d As New SqlCommand("select id from awm", con4)
            Dim reader3 As SqlDataReader
            reader3 = cmd1d.ExecuteReader

            While reader3.Read
                Dim sid = reader3.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader3.Close()

        ElseIf cb1.Text = "Mason" Then
            Dim cmd1d As New SqlCommand("select id from am", con4)
            Dim reader4 As SqlDataReader
            reader4 = cmd1d.ExecuteReader

            While reader4.Read
                Dim sid = reader4.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader4.Close()

        ElseIf cb1.Text = "Gardener" Then
            Dim cmd1d As New SqlCommand("select id from ag", con4)
            Dim reader5 As SqlDataReader
            reader5 = cmd1d.ExecuteReader

            While reader5.Read
                Dim sid = reader5.GetInt32(0)
                cb2.Items.Add(sid)
            End While

            reader5.Close()
        End If
    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        Call Button4_Click(sender, e)
    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con4.Open()
        cb1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class