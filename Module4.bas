Attribute VB_Name = "Module4"
Sub connPS()

Dim psM, nrM, psT, psM2, nrT, psM3, nrT3, nrTlT, g, h, nazt As String

st = ""

'x = psM y=nrM a=psT c=nrT b=psM2 d=psM3 e=nrT3 f= nrTlT

'Get number and ps form sheet "test"
For i = 2 To Sheets("test").Range("A" & Rows.Count).End(xlUp).Row
psM = Sheets("test").Cells(i, 2)
nrM = Sheets("test").Cells(i, 3)
    
    'Get only 1 row with PS
    If InStr(st, psM) <> 0 Then
        GoTo 200
    End If
    'Skip empty rows
    If psM = "" Then
        GoTo 200
    End If
    
    'String with checked ps
    st = st + psM
    
    'Get ps and name from 2nd list
    For j = 3 To Sheets("ltom").Range("A" & Rows.Count).End(xlUp).Row
    
        psT = Sheets("ltom").Cells(j, 728)
        nazt = Sheets("ltom").Cells(j, 4)
        'Skip non value rows
        If psT = "" Then
            GoTo 100
        ElseIf psT = "not found" Then
            GoTo 100
        ElseIf psT = 0 Then
            GoTo 100
        End If
    
    'Connect PSs from lists
    If InStr(psT, psM) = 0 Then
        GoTo 100
    ElseIf InStr(psT, psM) <> 0 Then
        
        For k = 2 To Sheets("test").Range("A" & Rows.Count).End(xlUp).Row
            psM2 = Sheets("test").Cells(k, 2)
            nrT = Sheets("test").Cells(k, 7)
                'If the same ps in lists and empty cell in 7 column add number
                If psM = psM2 And nrT = "" Then
                    Sheets("test").Cells(k, 7) = Sheets("ltom").Cells(j, 2)
                    Sheets("test").Cells(k, 10) = nazt
                    GoTo 100
                'If the same ps in lists and no empty cell in 7 column:
                ElseIf psM = psM2 And nrT <> 0 Then
                    For l = 2 To Sheets("test").Range("A" & Rows.Count).End(xlUp).Row
                        psM3 = Sheets("test").Cells(l, 2)
                        nrT3 = Sheets("test").Cells(l, 7)
                            If psM3 = psM And nrT3 = "" Then
                                Sheets("test").Cells(l, 7) = Sheets("ltom").Cells(j, 2)
                                Sheets("test").Cells(l, 10) = nazt
                                GoTo 100
                            End If
                    Next l
                'add number to n = first empty row
                n = Sheets("test").Range("A" & Rows.Count).End(xlUp).Row + 1
                Sheets("test").Cells(n, 1) = Sheets("test").Cells(i, 1)
                Sheets("test").Cells(n, 2) = psM
                Sheets("test").Cells(n, 11) = Sheets("test").Cells(i, 11)
                Sheets("test").Cells(n, 12) = Sheets("test").Cells(i, 12)
                Sheets("test").Cells(n, 13) = Sheets("test").Cells(i, 13)
                Sheets("test").Cells(n, 7) = Sheets("ltom").Cells(j, 2)
                Sheets("test").Cells(n, 10) = Sheets("ltom").Cells(j, 4)
                GoTo 100
            End If
        Next k
    End If
100 Next j

200 Next i

End Sub
