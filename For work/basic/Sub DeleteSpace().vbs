Sub deletespace()
    With ActiveSheet
        ROW1 = InputBox('93Row Number\'94)
        COLUMN1 = InputBox('93Column Number'94)
        For i = ROW1 To 1 Step -1
            For j = 1 To COLUMN1
                If Len(.Cells(i, j)) = 0 Then
                    .Cells(i, j).EntireRow.Delete
                    Exit For
                End If
            Next j
        Next i
    End With
End Sub