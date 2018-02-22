
Function MakeSQLListMaster(c As Range, Optional boundary As Integer = 1, Optional sep As Integer = 1)

    Dim str As String, d As Range, col As Range
    Dim sepchar As String, boundarychar As String

    Select Case boundary
        Case 1
            boundarychar = ""
        Case 2
            boundarychar = "'"
        Case 3
            boundarychar = """"
        Case 4
            boundarychar = "''"
    End Select
    
    Select Case sep
        Case 1
            sepchar = ","
        Case 2
            sepchar = ";"
    End Select
    
    For Each col In c.Columns
        For Each d In col.Cells
            If Not Trim(d) = "" Then
                str = str & boundarychar & Trim(d) & boundarychar
                If Not (d.row = (c.Cells(1, 1).row + c.Rows.Count - 1) And d.Column = (c.Cells(1, 1).Column + c.Columns.Count - 1)) Then
                    str = str & sepchar
                End If
            End If
        Next
    Next
    
    MakeSQLListMaster = str

End Function

