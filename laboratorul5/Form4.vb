Imports System.Windows.Forms.DataVisualization.Charting

Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ox() As Double
        Dim pres() As Double
        Dim ox2() As Double
        Dim pres2() As Double
        Dim peretex(2), peretey(2)
        ReDim ox(0 To Form1.straturi + 2)
        ReDim pres(0 To Form1.straturi + 2)
        ReDim ox2(0 To Form1.straturi + 3)
        ReDim pres2(0 To Form1.straturi + 3)

        For i = 0 To Form1.straturi + 2
            If i = 0 Then
                ox(i) = 0
            ElseIf i = 1 Then
                ox(i) = 50
            ElseIf i = Form1.straturi + 2 Then
                ox(i) = ox(i - 1) + 100
            Else ox(i) = ox(i - 1) + Form1.lungime(i - 2)
            End If
        Next

        For i = 0 To Form1.straturi + 2
            If i = 0 Or i = 1 Then
                pres(i) = Form1.pi
            ElseIf i = (Form1.straturi + 2) Or i = (Form1.straturi + 1) Then
                pres(i) = Form1.pe
            Else
                pres(i) = Form1.presiune_strat(i - 2)
            End If
        Next

        For i = 0 To Form1.straturi + 3
            If i = 0 Then
                ox2(i) = 0
            ElseIf i = 1 Then
                ox2(i) = 40
            ElseIf i = 2 Then
                ox2(i) = 50
            ElseIf i = Form1.straturi + 3 Then
                ox2(i) = ox2(i - 1) + 100
            Else ox2(i) = ox2(i - 1) + Form1.lungime(i - 3)
            End If
        Next

        For i = 0 To Form1.straturi + 3
            If i = 0 Or i = 1 Then
                pres2(i) = Form1.pisat
            ElseIf i = (Form1.straturi + 3) Then
                pres2(i) = Form1.pesat
            Else
                pres2(i) = Form1.presiune_sat(i - 2)
            End If
        Next

        Chart2.Series.Add("Presiunea_Vap")
        Chart2.Series("Presiunea_Vap").ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series("Presiunea_Vap").Points.DataBindXY(ox, pres)

        Chart2.Series.Add("Presiunea_sat_Vap")
        Chart2.Series("Presiunea_sat_Vap").ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series("Presiunea_sat_Vap").Points.DataBindXY(ox2, pres2)

        For i = 1 To Form1.straturi + 1
            peretex(0) = ox(i)
            peretex(1) = ox(i)
            peretey(0) = pres2(0) + 200
            peretey(1) = 0
            If i = 1 Then
                Chart2.Series.Add("Perete_interior")
                Chart2.Series("Perete_interior").ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart2.Series("Perete_interior").Points.DataBindXY(peretex, peretey)
            ElseIf i = Form1.straturi + 1 Then
                Chart2.Series.Add("Perete_exterior")
                Chart2.Series("Perete_exterior").ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart2.Series("Perete_exterior").Points.DataBindXY(peretex, peretey)
            Else
                Dim peretenou As String = ("Strat" & i)
                Chart2.Series.Add(peretenou)
                Chart2.Series(peretenou).ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart2.Series(peretenou).Points.DataBindXY(peretex, peretey)
            End If

        Next


    End Sub
End Class