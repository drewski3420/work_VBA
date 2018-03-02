Option Explicit

Sub SingleFolder()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim fold_txt As String, ssrs_mapped_folder As String, command_text As String, textRow As String
    Dim f As file, subfold As Folder, fso As New FileSystemObject, fold As Folder
    Dim fileNo As Long, is_stored_proc As Long, i As Long, j As Long

    Set ws = ActiveSheet

    On Error Resume Next
    ssrs_mapped_folder = "C:\Users\andrewj\SSRS Mapped Folder\Selene Reporting"
    Set fold = fso.GetFolder(ssrs_mapped_folder)
    If Err.Number <> 0 Then
        MsgBox "SSRS Mapped Folder not found. Change the ssrs_mapped_folder variable to match your system.", vbExclamation + vbOKOnly + vbCritical, "Error!"
        Exit Sub
    End If
    fold_txt = InputBox("Which Folder?", "Choose a Folder")
    Set subfold = fso.GetFolder(fso.BuildPath(ssrs_mapped_folder, fold_txt))
    If Err.Number <> 0 Then
        MsgBox "Subfolder " & fold_txt & " not found. Check the folder name and try again.", vbExclamation + vbOKOnly + vbCritical, "Error!"
        Exit Sub
    End If
    On Error GoTo 0
    
    ws.UsedRange.ClearContents

    i = 0
    For Each f In subfold.Files
        DoEvents
        If f.Type = "RDL File" Then
            fileNo = FreeFile
            command_text = ""
            j = 0
            ws.Cells(2, 1).Offset(i, 0) = subfold.Name
            ws.Cells(2, 1).Offset(i, 1) = f.Name
            Open f.Path For Input As #fileNo
            Do While Not EOF(fileNo)
                DoEvents
                Line Input #fileNo, textRow
                If textRow Like "*<CommandType>StoredProcedure</CommandType>*" Then
                    is_stored_proc = 1
                End If
                If textRow Like "*<CommandText>*" And is_stored_proc = 1 Then
                    command_text = Trim(Replace(Replace(textRow, "<CommandText>", ""), "</CommandText>", ""))
                    ws.Cells(2, 1).Offset(i, 2 + j) = command_text
                    j = j + 1
                    is_stored_proc = 0
                End If
                
            Loop
            Close #fileNo
            i = i + 1
        End If
    Next
    RemoveDupes
    Consolidate
    
    ws.Range("A1") = "Folder"
    ws.Range("B1") = "Report Name"
    For i = 3 To ws.Range("A1").SpecialCells(xlCellTypeLastCell).Column
        DoEvents
        ws.Cells(1, i) = "Proc" & (i - 2)
    Next
    Application.ScreenUpdating = True
End Sub
Sub RemoveDupes()
    Dim lastCol As Integer, lastRow As Integer
    
    Dim ws As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim curr As String, comp As String
    Set ws = ActiveSheet

    lastCol = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Column
    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    For i = lastCol To 3 Step -1
        DoEvents
        For j = 2 To lastRow Step 1
            DoEvents
            curr = ws.Cells(j, i)
            If curr <> "" Then
                For k = i - 1 To 3 Step -1
                    DoEvents
                    comp = ws.Cells(j, k)
                    If curr = comp Then
                        ws.Cells(j, i) = ""
                        Exit For
                    End If
                Next
            End If
        Next
    Next
End Sub

Sub Consolidate()
    Dim lastCol As Integer, lastRow As Integer
    
    Dim ws As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim curr As String, comp As String
    Set ws = ActiveSheet

    lastCol = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Column
    lastRow = ws.Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    For i = lastCol To 3 Step -1
        DoEvents
        For j = 2 To lastRow Step 1
            DoEvents
            curr = ws.Cells(j, i)
            If curr <> "" Then
                For k = 3 To i - 1 Step 1
                    DoEvents
                    comp = ws.Cells(j, k)
                    If comp = "" Then
                        ws.Cells(j, k) = curr
                        ws.Cells(j, i) = ""
                        Exit For
                    End If
                Next
            End If
        Next
    Next
End Sub
