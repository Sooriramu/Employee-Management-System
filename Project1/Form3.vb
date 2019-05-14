Imports System.Data.SqlClient
Imports System.Data

Public Class Form3
    Dim con2 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")
    Dim command As New SqlCommand
    Dim n1 As Integer

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con2.Open()
        cb1.SelectedIndex = 0
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If cb1.Text = "Men mazdoor" Then
            If TextBox1.Text <> "" Then
                Dim cmd3c As New SqlCommand("select max(id) FROM amm ", con2)
                Dim dr As SqlDataReader
                dr = cmd3c.ExecuteReader
                dr.Read()
                TextBox2.Text = dr.GetValue(0)
                n1 = TextBox2.Text
                n1 = n1 + 1
                TextBox2.Text = n1
                dr.Close()
                Dim cmd3a As New SqlCommand("insert into amm(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox1.Text + "')  ", con2)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("insert into otmm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("insert into ltmm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Inserted")

                Me.Close()

            Else
                MessageBox.Show("Enter the name")
            End If

        ElseIf cb1.Text = "Carpenter" Then
            If TextBox1.Text <> "" Then
                Dim cmd3c As New SqlCommand("select max(id) FROM ac ", con2)
                Dim dr As SqlDataReader
                dr = cmd3c.ExecuteReader
                dr.Read()
                TextBox2.Text = dr.GetValue(0)
                n1 = TextBox2.Text
                n1 = n1 + 1
                TextBox2.Text = n1
                dr.Close()
                Dim cmd3a As New SqlCommand("insert into ac(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox1.Text + "')  ", con2)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("insert into otc (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("insert into ltc (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Inserted")

                Me.Close()

            Else
                MessageBox.Show("Enter the name")
            End If

        ElseIf cb1.Text = "Women mazdoor" Then
            If TextBox1.Text <> "" Then
                Dim cmd3c As New SqlCommand("select max(id) FROM awm ", con2)
                Dim dr As SqlDataReader
                dr = cmd3c.ExecuteReader
                dr.Read()
                TextBox2.Text = dr.GetValue(0)
                n1 = TextBox2.Text
                n1 = n1 + 1
                TextBox2.Text = n1
                dr.Close()
                Dim cmd3a As New SqlCommand("insert into awm(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox1.Text + "')  ", con2)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("insert into otwm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("insert into ltwm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Inserted")

                Me.Close()

            Else
                MessageBox.Show("Enter the name")
            End If

        ElseIf cb1.Text = "Mason" Then
            If TextBox1.Text <> "" Then
                Dim cmd3c As New SqlCommand("select max(id) FROM am ", con2)
                Dim dr As SqlDataReader
                dr = cmd3c.ExecuteReader
                dr.Read()
                TextBox2.Text = dr.GetValue(0)
                n1 = TextBox2.Text
                n1 = n1 + 1
                TextBox2.Text = n1
                dr.Close()
                Dim cmd3a As New SqlCommand("insert into am(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox1.Text + "')  ", con2)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("insert into otm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("insert into ltm (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Inserted")

                Me.Close()

            Else
                MessageBox.Show("Enter the name")
            End If

        ElseIf cb1.Text = "Gardener" Then
            If TextBox1.Text <> "" Then
                Dim cmd3c As New SqlCommand("select max(id) FROM ag ", con2)
                Dim dr As SqlDataReader
                dr = cmd3c.ExecuteReader
                dr.Read()
                TextBox2.Text = dr.GetValue(0)
                n1 = TextBox2.Text
                n1 = n1 + 1
                TextBox2.Text = n1
                dr.Close()
                Dim cmd3a As New SqlCommand("insert into ag(ID,Name) values('" + TextBox2.Text + "' ,'" + TextBox1.Text + "')  ", con2)
                cmd3a.ExecuteNonQuery()

                Dim cmd3d As New SqlCommand("insert into otg (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3d.ExecuteNonQuery()

                Dim cmd3e As New SqlCommand("insert into ltg (Name,ID) values('" + TextBox1.Text + "','" + TextBox2.Text + "')", con2)
                cmd3e.ExecuteNonQuery()

                MessageBox.Show("Inserted")

                Me.Close()

            Else
                MessageBox.Show("Enter the name")
            End If

        End If
        TextBox1.Text = ""

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        TextBox1.Text = ""
        Me.Close()

    End Sub

    Private Sub cb6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        TextBox1.Text = ""
        Dim cmd3b As New SqlDataAdapter
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
End Class