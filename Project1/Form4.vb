Imports System.Data
Imports System.Data.SqlClient
Public Class Form4
    Dim con3 As New SqlConnection("Data Source=SOORI\SQLEXPRESS;Initial Catalog=proj1;Integrated Security=True")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            If TextBox1.Text = TextBox2.Text Then
                Dim cmd4a As New SqlCommand("update login set pin='" + TextBox1.Text + "' where username='psg'", con3)
                cmd4a.ExecuteNonQuery()
                MessageBox.Show("Pin updated successfully")
                TextBox1.Text = ""
                TextBox2.Text = ""
                Me.Hide()
                Form2.Show()

            Else
                MessageBox.Show("Pin mismatching")
            End If
        Else
            MessageBox.Show("Textfields cannot be empty")
        End If
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con3.Open()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        Me.Hide()
        Form2.Show()
    End Sub
End Class