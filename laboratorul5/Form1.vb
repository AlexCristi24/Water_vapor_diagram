
Imports System.Data.SqlClient
Imports System.Windows
Imports System.Math
Imports System.Runtime.Intrinsics.X86
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1

    Public temp_int,temp_ext,umid_int,umid_ext,lungime(), presiune_strat(), rez_temp(), temp(), presiune_sat(),miu_coef(),lambda_coef(),miu() As Double
    Public straturi As Integer
    Public materiale(), nume(), material() As String
    Public valori As Integer
    Public lambda(), rvp(), pi, pe, RVT, RT, pisat, pesat As Double

    'clasa pentru afisarea bazei de date
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connection As SqlConnection = New SqlConnection()
        connection.ConnectionString = "Data Source=DESKTOP-AN74KRO;
         Initial Catalog=Valori_Materiale;Integrated Security=True"
        connection.Open()
        Dim adp As SqlDataAdapter = New SqlDataAdapter _
        ("select * from Sheet1", connection)
        Dim ds As DataSet = New DataSet()
        adp.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        valori = (DataGridView1.RowCount - 2)
    End Sub
    'clasa pentru citirea din baza de date a tuturor informatiilor
    Public Sub citim_date()
        For i = 0 To valori
            lambda(i) = DataGridView1.Rows(i).Cells(2).Value
            miu(i) = DataGridView1.Rows(i).Cells(3).Value
            nume(i) = DataGridView1.Rows(i).Cells(0).Value
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Show()
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Calcul_valori()
    End Sub

    'clasa pentru citirea parametrilor de la tastatura
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        temp_int = TextBox1.Text
        temp_ext = TextBox2.Text
        umid_int = TextBox5.Text
        umid_ext = TextBox4.Text
        straturi = TextBox3.Text
        ReDim nume(0 To valori)
        ReDim lambda(0 To valori)
        ReDim miu(0 To valori)
        ReDim material(0 To straturi - 1)
        ReDim lambda_coef(0 To straturi - 1)
        ReDim miu_coef(0 To straturi - 1)
        ReDim lungime(0 To straturi - 1)
        ReDim rvp(0 To straturi - 1)
        ReDim presiune_strat(0 To straturi - 2)
        ReDim rez_temp(0 To straturi + 1)
        ReDim temp(0 To straturi)
        ReDim presiune_sat(0 To straturi)
        citim_date()
        Form3.Show()
    End Sub
    Function calcul_pres_sat(temp As Double) As Double
        Dim ps1 As Double
        If temp > 0 Then
            Dim x As Double = (17.269 * temp / (237.3 + temp))
            ps1 = 610.5 * Pow(Math.E, x)
        Else
            Dim x As Double = (21.875 * temp / (265.5 + temp))
            ps1 = 610.5 * Pow(Math.E, x)
        End If
        Return ps1
    End Function

    Function calcul_pres(pres_sat As Double, umid As Double) As Double
        Return pres_sat * umid / 100
    End Function

    Function calcul_rez_vap(dist As Double, miu As Double) As Double
        Return 50 * Pow(10, 8) * dist / 1000 * miu
    End Function

    Function calcul_presiune_strat(o As Integer) As Double
        Dim rvpart As Double
        rvpart = 0
        For i = 0 To o
            rvpart = rvp(i) + rvpart
        Next
        Return (pi - (rvpart / RVT * (pi - pe)))
    End Function

    Function calcul_temp(o As Integer) As Double
        Dim rpart As Double
        rpart = 0
        For i = 0 To o
            rpart = rez_temp(i) + rpart
        Next
        Return temp_int - rpart / RT * (temp_int - temp_ext)
    End Function

    Function calcul_rez_temp(o As Integer) As Double
        If lambda_coef(o) = 0 Then
            Return 0
        Else
            Return lungime(o) / lambda_coef(o) / 1000
        End If
    End Function

    Private Sub Calcul_valori()
        pi = calcul_pres(calcul_pres_sat(temp_int), umid_int)
        pisat = calcul_pres_sat(temp_int)
        pesat = calcul_pres_sat(temp_ext)
        pe = calcul_pres(calcul_pres_sat(temp_ext), umid_ext)
        RVT = 0
        RT = 0
        For i = 0 To straturi - 1
            rvp(i) = calcul_rez_vap(lungime(i), miu_coef(i))
            RVT = RVT + rvp(i)
        Next
        For i = 0 To straturi - 2
            presiune_strat(i) = calcul_presiune_strat(i)
        Next
        For i = 0 To straturi + 1
            If i = 0 Then
                rez_temp(i) = 0.125
            ElseIf i = straturi + 1 Then
                rez_temp(i) = 0.042
            Else rez_temp(i) = calcul_rez_temp(i - 1)
            End If
            RT = RT + rez_temp(i)
        Next
        For i = 0 To straturi
            temp(i) = calcul_temp(i)
            presiune_sat(i) = calcul_pres_sat(temp(i))
        Next
    End Sub
End Class
