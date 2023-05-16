Attribute VB_Name = "Module2"
Function MODIFYTEXT(dataRange As Range, action As String, position As String, charNum As Long, newText As String) As Variant
    Dim i As Long
    Dim oldText As String
    Dim result() As Variant
    Dim cellValue As Variant
    
    ReDim result(1 To dataRange.Cells.Count, 1 To 1)
    
    For i = 1 To dataRange.Cells.Count
        cellValue = dataRange.Cells(i).Value
        If Not IsError(cellValue) Then
            oldText = CStr(cellValue)
            If action = "Add" Then
                If position = "Position" Then
                    result(i, 1) = Left(oldText, charNum) & newText & Right(oldText, Len(oldText) - charNum)
                Else
                    result(i, 1) = oldText & newText
                End If
            ElseIf action = "Remove" Then
                If position = "Position" Then
                    result(i, 1) = Left(oldText, charNum - 1) & Right(oldText, Len(oldText) - charNum - Len(newText) + 1)
                Else
                    result(i, 1) = Replace(oldText, newText, "")
                End If
            Else
                result(i, 1) = "Invalid action"
            End If
        Else
            result(i, 1) = cellValue
        End If
    Next i
    
    MODIFYTEXT = result
End Function

