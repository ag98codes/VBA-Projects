Attribute VB_Name = "Module1"
Function WEIGHTAVG(dataRange As Range, weightRange As Range) As Double
    Dim i As Long
    Dim sumData As Double
    Dim sumWeight As Double
    
    For i = 1 To dataRange.Cells.Count
        sumData = sumData + (dataRange.Cells(i).Value * weightRange.Cells(i).Value)
        sumWeight = sumWeight + weightRange.Cells(i).Value
    Next i

    WEIGHTAVG = sumData / sumWeight
End Function



