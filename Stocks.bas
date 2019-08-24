Attribute VB_Name = "Module11"
Sub stocks1():

Dim I As Long
Dim rowcounter As Integer
Dim Sum As Double
Dim lastrow As Long
Dim oprate As Double
Dim clrate As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percentage Change"

rowcounter = 2
Sum = 0
orate = Cells(2, 3).Value
clrate = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To lastrow
If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
Sum = Sum + Cells(I, 7).Value
clrate = Cells(I, 6).Value
            
' if they're different add ticker name and volumesum in column 9 and 10
Cells(rowcounter, 9).Value = Cells(I, 1).Value
Cells(rowcounter, 10).Value = Sum
Cells(rowcounter, 11).Value = clrate - orate
If Cells(rowcounter, 11).Value > 0 Then
Cells(rowcounter, 11).Interior.ColorIndex = 4
Else
Cells(rowcounter, 11).Interior.ColorIndex = 3
End If
If orate = 0 Then
    Cells(rowcounter, 12).Value = "Zero Open Rate"
    Else
    Cells(rowcounter, 12).Value = Cells(rowcounter, 11).Value / orate
    End If
Cells(rowcounter, 12).NumberFormat = "0.00%"
orate = Cells(I + 1, 3).Value
rowcounter = rowcounter + 1
Sum = 0
Else
Sum = Sum + Cells(I, 7).Value
End If
Next I
End Sub



