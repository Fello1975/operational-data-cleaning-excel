Sub FillBlanksAcrossColumns()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim ColIndex As Integer
    Dim i As Long

    ' Set the worksheet variable to the active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Loop through each column in the worksheet
    For ColIndex = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        LastRow = ws.Cells(ws.Rows.Count, ColIndex).End(xlUp).Row

        ' Start from the second row to the last row
        For i = 2 To LastRow
            If IsEmpty(ws.Cells(i, ColIndex).Value) And Not IsEmpty(ws.Cells(i - 1, ColIndex).Value) Then
                ' Fill down the value from the last non-empty cell above
                ws.Cells(i, ColIndex).Value = ws.Cells(i - 1, ColIndex).Value
                ' Continue filling down until a non-empty cell is found
                Do While IsEmpty(ws.Cells(i + 1, ColIndex).Value) And i < LastRow
                    ws.Cells(i + 1, ColIndex).Value = ws.Cells(i, ColIndex).Value
                    i = i + 1
                Loop
            End If
        Next i
    Next ColIndex
End Sub
