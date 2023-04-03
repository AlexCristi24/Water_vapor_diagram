Imports System.Data.SqlClient
Imports System.Linq.Expressions
Imports System.Reflection.Emit

Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = Format(Form1.pi, "#.00")
        Label4.Text = Format(Form1.pe, "#.00")
        Label11.Text = Format(Form1.RVT, "#.00")
        Label13.Text = Format(Form1.RT, "#.00")
        For i = 0 To Form1.straturi - 1
            ListBox1.Items.Add(Format(Form1.rvp(i), "#.00"))
        Next
        For i = 0 To Form1.straturi - 2
            ListBox2.Items.Add(Format(Form1.presiune_strat(i), "0.00"))
        Next
        For i = 1 To Form1.straturi
            ListBox3.Items.Add(Format(Form1.rez_temp(i), "0.00"))
        Next
        For i = 0 To Form1.straturi
            ListBox4.Items.Add(Format(Form1.temp(i), "0.00"))
            ListBox5.Items.Add(Format(Form1.presiune_sat(i), "0.00"))
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form4.Show()
    End Sub
End Class