Attribute VB_Name = "Module1"
Sub output()
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

Dim i As Double
Dim LastRow As Double
Dim VarCheck As String
Dim addo As Double
Dim addo2 As Double
Dim LOWVAL As Double
Dim HIGHVAL As Double
Dim VolAccum As Double
LOWVAL = 0
HIGHVAL = 0
LastRow = Range("A" & Rows.Count).End(xlUp).Row
addo = 2
addo2 = 2
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("J:L").AutoFit

For i = 2 To LastRow
VarCheck = Cells(i, 1).Value
If VarCheck = Cells(i + 1, 1).Value Then
    Cells(addo, 9).Value = VarCheck
Else
    addo = addo + 1
        
End If

If VarCheck <> Cells(i - 1, 1).Value Then
    LOWVAL = Cells(i, 3).Value
    addo2 = addo2 + 1
    Cells(addo2 - 1, 11).NumberFormat = "0.00%"
Else
    If LOWVAL = 0 Then
        Cells(addo2 - 1, 10) = 0
        Cells(addo2 - 1, 11) = 0
        Cells(addo2 - 1, 12) = 0
    Else
        HIGHVAL = Cells(i, 6).Value
        Cells(addo2 - 1, 10) = HIGHVAL - LOWVAL
        comb = Cells(addo2 - 1, 10)
        Cells(addo2 - 1, 11) = ((HIGHVAL - LOWVAL) / LOWVAL)
    End If
End If

If Cells(i, 7).Value = 0 Then
    VolAccum = 0
Else
    VolAccum = VolAccum + Cells(i, 7).Value
    Cells(addo2 - 1, 12).Value = VolAccum
End If

Next i
For i = 2 To LastRow
If Cells(i, 10) < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
ElseIf Cells(i, 10) > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
End If
Next i

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


Dim HIGHEST As Double
Dim LOWEST As Double
Dim HIGHEST_TICKER As String
Dim LOWEST_TICKER As String
Dim VOL As Double
Dim VOL_TICKER As String
HIGHEST = 0
LOWEST = 0
VOL = 0

For i = 2 To LastRow
If Cells(i, 11).Value > HIGHEST Then
HIGHEST = Cells(i, 11).Value
HIGHEST_TICKER = Cells(i, 9).Value

End If

If Cells(i, 11).Value < LOWEST Then
LOWEST = Cells(i, 11).Value
LOWEST_TICKER = Cells(i, 9).Value

End If

If Cells(i, 12).Value > VOL Then
VOL = Cells(i, 12).Value
VOL_TICKER = Cells(i, 9).Value

End If
Next i


Range("Q2").Value = HIGHEST
Range("Q3").Value = LOWEST
Range("P2").Value = HIGHEST_TICKER
Range("P3").Value = LOWEST_TICKER
Range("Q4").Value = VOL
Range("P4").Value = VOL_TICKER
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"
Columns("O:Q").AutoFit

Next

starting_ws.Activate
End Sub

