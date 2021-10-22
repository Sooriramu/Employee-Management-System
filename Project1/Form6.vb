Imports System.Data
Imports System.Data.SqlClient
Public Class Form6
    Dim con5 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2)
        con5.Open()
        cb1.SelectedIndex = 0
        Label5.Text = ""
    End Sub

    Private Sub cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged

        cb2.Items.Clear()
        Label5.Text = ""
        TextBox1.Text = ""

        Try

            If cb1.Text = "Men mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from amm", con5)
                Dim reader1 As SqlDataReader
                reader1 = cmd1d.ExecuteReader

                While reader1.Read
                    Dim sid = reader1.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader1.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1d As New SqlCommand("select id from ac", con5)
                Dim reader2 As SqlDataReader
                reader2 = cmd1d.ExecuteReader

                While reader2.Read
                    Dim sid = reader2.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader2.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1d As New SqlCommand("select id from awm", con5)
                Dim reader3 As SqlDataReader
                reader3 = cmd1d.ExecuteReader

                While reader3.Read
                    Dim sid = reader3.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader3.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1d As New SqlCommand("select id from am", con5)
                Dim reader4 As SqlDataReader
                reader4 = cmd1d.ExecuteReader

                While reader4.Read
                    Dim sid = reader4.GetString(0)
                    cb2.Items.Add(sid)
                End While
                reader4.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1d As New SqlCommand("select id from ag", con5)
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
        TextBox1.Text = ""
        Try
            If cb1.Text = "Men mazdoor" Then
                Dim cmd1y As New SqlCommand("select Name from amm where id='" + cb2.Text + "'", con5)
                Dim dr As SqlDataReader
                dr = cmd1y.ExecuteReader
                dr.Read()
                Label5.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Carpenter" Then
                Dim cmd1z As New SqlCommand("select Name from ac where id='" + cb2.Text + "'", con5)
                Dim dr As SqlDataReader
                dr = cmd1z.ExecuteReader
                dr.Read()
                Label5.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Women mazdoor" Then
                Dim cmd1aa As New SqlCommand("select Name from awm where id='" + cb2.Text + "'", con5)
                Dim dr As SqlDataReader
                dr = cmd1aa.ExecuteReader
                dr.Read()
                Label5.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Mason" Then
                Dim cmd1ab As New SqlCommand("select Name from am where id='" + cb2.Text + "'", con5)
                Dim dr As SqlDataReader
                dr = cmd1ab.ExecuteReader
                dr.Read()
                Label5.Text = dr.GetValue(0)
                dr.Close()
            ElseIf cb1.Text = "Gardener" Then
                Dim cmd1ac As New SqlCommand("select Name from ag where id='" + cb2.Text + "'", con5)
                Dim dr As SqlDataReader
                dr = cmd1ac.ExecuteReader
                dr.Read()
                Label5.Text = dr.GetValue(0)
                dr.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If cb2.Text <> "" Then
            If TextBox1.Text <> "" Then
                TextBox1.Text = StrConv(TextBox1.Text, vbProperCase)
                TextBox3.Text = TextBox1.Text + " " + TextBox2.Text

                Try

                    If cb1.Text = "Men mazdoor" Then


                        Dim cmd3a As New SqlCommand("update amm set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("update otltmm set name='" + TextBox3.Text + "'where id='" + cb2.Text + "'  ", con5)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Renamed successfully")


                    ElseIf cb1.Text = "Carpenter" Then

                        Dim cmd3a As New SqlCommand("update ac set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("update otltc set name='" + TextBox3.Text + "'where id='" + cb2.Text + "'  ", con5)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Renamed successfully")

                    ElseIf cb1.Text = "Women mazdoor" Then

                        Dim cmd3a As New SqlCommand("update awm set name='" + TextBox3.Text + "'where id='" + cb2.Text + "'  ", con5)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("update otltwm set name='" + TextBox3.Text + "'where id='" + cb2.Text + "'  ", con5)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Renamed successfully")

                    ElseIf cb1.Text = "Mason" Then

                        Dim cmd3a As New SqlCommand("update am set name='" + TextBox3.Text + "'where id='" + cb2.Text + "'  ", con5)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("update otltm set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Renamed successfully")

                    ElseIf cb1.Text = "Gardener" Then

                        Dim cmd3a As New SqlCommand("update ag set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                        cmd3a.ExecuteNonQuery()

                        Dim cmd3d As New SqlCommand("update otltg set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                        cmd3d.ExecuteNonQuery()

                        MessageBox.Show("Renamed successfully")

                    End If

                    Dim cmd3e As New SqlCommand("update register set name='" + TextBox3.Text + "'where id='" + cb2.Text + "' ", con5)
                    cmd3e.ExecuteNonQuery()

                    Me.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Else
                MessageBox.Show("Name cannot be empty")
            End If
        Else
            MessageBox.Show("Select the id")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form6_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

        Try
            If TextBox1.Text = "" Then
                Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890 "
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If
            Else
                Dim dac As String = ",<>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890"
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

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        Try

            Dim dac As String = ",<.>/?;:'[{]}=+-_()*&^%$#@!~`""1234567890 "
                If dac.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                    MessageBox.Show("Only alphabets are allowed")
                End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class