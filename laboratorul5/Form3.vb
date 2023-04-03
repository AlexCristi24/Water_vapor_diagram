Imports System.Security.Cryptography.X509Certificates
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim straturi2 = Form1.straturi
        Dim lstb As Forms.ListBox
        Dim lbl As Forms.Label
        Dim tb As Forms.TextBox
        Dim lbl2 As Forms.Label
        For o = 1 To straturi2
            lstb = Controls("ListBox" & o)
            lbl = Controls("Label" & o)
            tb = Controls("TextBox" & o)
            lbl2 = Controls("Label" & (o + 9))
            lstb.Visible = True
            lbl.Visible = True
            tb.Visible = True
            lbl2.Visible = True
        Next
        For i = 0 To Form1.valori
            ListBox1.Items.Add(Form1.nume(i))
            ListBox2.Items.Add(Form1.nume(i))
            ListBox3.Items.Add(Form1.nume(i))
            ListBox4.Items.Add(Form1.nume(i))
            ListBox5.Items.Add(Form1.nume(i))
            ListBox6.Items.Add(Form1.nume(i))
            ListBox7.Items.Add(Form1.nume(i))
            ListBox8.Items.Add(Form1.nume(i))
            ListBox9.Items.Add(Form1.nume(i))
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Form1.material(0) = ListBox1.SelectedItem
        Form1.lambda_coef(0) = Form1.lambda(ListBox1.SelectedIndex)
        Form1.miu_coef(0) = Form1.miu(ListBox1.SelectedIndex)
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) Then
            Form1.lungime(0) = TextBox1.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Form1.material(1) = ListBox2.SelectedItem
        Form1.lambda_coef(1) = Form1.lambda(ListBox2.SelectedIndex)
        Form1.miu_coef(1) = Form1.miu(ListBox2.SelectedIndex)
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If IsNumeric(TextBox2.Text) Then
            Form1.lungime(1) = TextBox2.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Form1.material(2) = ListBox3.SelectedItem
        Form1.lambda_coef(2) = Form1.lambda(ListBox3.SelectedIndex)
        Form1.miu_coef(2) = Form1.miu(ListBox3.SelectedIndex)
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) Then
            Form1.lungime(2) = TextBox3.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        Form1.material(3) = ListBox4.SelectedItem
        Form1.lambda_coef(3) = Form1.lambda(ListBox4.SelectedIndex)
        Form1.miu_coef(3) = Form1.miu(ListBox4.SelectedIndex)
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) Then
            Form1.lungime(3) = TextBox4.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox5.SelectedIndexChanged
        Form1.material(4) = ListBox5.SelectedItem
        Form1.lambda_coef(4) = Form1.lambda(ListBox5.SelectedIndex)
        Form1.miu_coef(4) = Form1.miu(ListBox5.SelectedIndex)
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) Then
            Form1.lungime(4) = TextBox5.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox6.SelectedIndexChanged
        Form1.material(5) = ListBox6.SelectedItem
        Form1.lambda_coef(5) = Form1.lambda(ListBox6.SelectedIndex)
        Form1.miu_coef(5) = Form1.miu(ListBox6.SelectedIndex)
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox6.Text) Then
            Form1.lungime(5) = TextBox6.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox7.SelectedIndexChanged
        Form1.material(6) = ListBox7.SelectedItem
        Form1.lambda_coef(6) = Form1.lambda(ListBox7.SelectedIndex)
        Form1.miu_coef(6) = Form1.miu(ListBox7.SelectedIndex)
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If IsNumeric(TextBox7.Text) Then
            Form1.lungime(6) = TextBox7.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox8.SelectedIndexChanged
        Form1.material(7) = ListBox8.SelectedItem
        Form1.lambda_coef(7) = Form1.lambda(ListBox8.SelectedIndex)
        Form1.miu_coef(7) = Form1.miu(ListBox8.SelectedIndex)
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If IsNumeric(TextBox8.Text) Then
            Form1.lungime(7) = TextBox8.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
    Private Sub ListBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox9.SelectedIndexChanged
        Form1.material(8) = ListBox9.SelectedItem
        Form1.lambda_coef(8) = Form1.lambda(ListBox9.SelectedIndex)
        Form1.miu_coef(8) = Form1.miu(ListBox9.SelectedIndex)
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) Then
            Form1.lungime(8) = TextBox9.Text
        Else
            MsgBox("Va rog introduceti un numar!")
        End If
    End Sub
End Class