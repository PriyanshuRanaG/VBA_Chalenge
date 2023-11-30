Attribute VB_Name = "Module1"
Sub AssignmentForAllSheets()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Call Assignment2(ws)
    Next ws
End Sub

Sub Assignment2(ws1 As Worksheet)
    Dim OpeningValue As Double
    Dim IndexRow As Long, i As Long
    Dim Total_Volume As Double
    Dim MaxIncrease As Double, MaxDecrease As Double, MaxVolume As Double
    Dim MaxIncreaseTicker As String, MaxDecreaseTicker As String, MaxVolumeTicker As String

    ' Initialize the maximum and minimum values
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0

    'Headers
    ws1.Range("I1").Value = "Ticker"
    ws1.Range("J1").Value = "Yearly Change"
    ws1.Range("K1").Value = "Percentage Change"
    ws1.Range("L1").Value = "Total Stock Volume"

    IndexRow = 2
    OpeningValue = ws1.Cells(2, 3).Value
    Total_Volume = 0

    For i = 2 To ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
        Total_Volume = Total_Volume + ws1.Cells(i, 7).Value

        If ws1.Cells(i + 1, 1).Value <> ws1.Cells(i, 1).Value Then
            Dim YearlyChange As Double
            YearlyChange = ws1.Cells(i, 6).Value - OpeningValue
            ws1.Cells(IndexRow, 9).Value = ws1.Cells(i, 1).Value
            ws1.Cells(IndexRow, 10).Value = YearlyChange
            ws1.Cells(IndexRow, 12).Value = Total_Volume

            Dim PercentChange As Double
            If OpeningValue <> 0 Then
                PercentChange = (YearlyChange / OpeningValue) * 100
                ws1.Cells(IndexRow, 11).Value = Round(PercentChange, 2)
            Else
                PercentChange = 0
                ws1.Cells(IndexRow, 11).Value = 0
            End If

            ' Check for max and min percentage change and max volume
            If PercentChange > MaxIncrease Then
                MaxIncrease = PercentChange
                MaxIncreaseTicker = ws1.Cells(i, 1).Value
            End If

            If PercentChange < MaxDecrease Then
                MaxDecrease = PercentChange
                MaxDecreaseTicker = ws1.Cells(i, 1).Value
            End If

            If Total_Volume > MaxVolume Then
                MaxVolume = Total_Volume
                MaxVolumeTicker = ws1.Cells(i, 1).Value
            End If

            IndexRow = IndexRow + 1
            If ws1.Cells(i + 1, 3).Value <> "" Then
                OpeningValue = ws1.Cells(i + 1, 3).Value
            End If
            Total_Volume = 0
        End If

        ' Change color based on positive or negative yearly change
        If ws1.Cells(IndexRow - 1, 10).Value <= 0 Then
            ws1.Cells(IndexRow - 1, 10).Interior.ColorIndex = 3
        ElseIf ws1.Cells(IndexRow - 1, 10).Value > 0 Then
            ws1.Cells(IndexRow - 1, 10).Interior.ColorIndex = 4
        End If

    Next i

    ' Output the greatest % increase, decrease, and total volume
    ws1.Range("O2").Value = "Greatest % Increase"
    ws1.Range("P2").Value = MaxIncreaseTicker
    ws1.Range("Q2").Value = MaxIncrease

    ws1.Range("O3").Value = "Greatest % Decrease"
    ws1.Range("P3").Value = MaxDecreaseTicker
    ws1.Range("Q3").Value = MaxDecrease

    ws1.Range("O4").Value = "Greatest Total Volume"
    ws1.Range("P4").Value = MaxVolumeTicker
    ws1.Range("Q4").Value = MaxVolume
End Sub


