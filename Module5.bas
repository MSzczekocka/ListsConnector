Attribute VB_Name = "Module5"
Sub markParts()

Dim nrM, nrT As String

' Get number from list 1
For i = 2 To Sheets("test").Range("A" & Rows.Count).End(xlUp).Row
    nrM = Sheets("test").Cells(i, 3)
    'Skip empty cells
    If nrM = "" Then
        GoTo 100
    End If
    
    'Get number from list 2
    For j = 2 To Sheets("test").Range("A" & Rows.Count).End(xlUp).Row
        nrT = Sheets("test").Cells(j, 7)
        'Skip empty cells
        If nrT = "" Then
            GoTo 50
        End If
        
        'Mark the same numbers
        If nrM = nrT Then
            Sheets("test").Cells(j, 7).Interior.Color = RGB(189, 215, 238)
        End If
50    Next j
100 Next i

End Sub
