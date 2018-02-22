Option Explicit
Dim pubCnt As Integer 'used to see if any of the strip functions did anything

Function StripAll(str As String) As String
    
    Do While True
        DoEvents
        pubCnt = 0
        str = StripDisclaimers(str) 'strip disclaimers
        str = StripSomeThings(str) 'do all the stripping first
        str = StripSingleSpaceLines(str) 'then the lines
        str = StripMultipleLines(str) 'replace multiple lines with only 1
        
        If pubCnt = 0 Then: Exit Do
        
    Loop
    
    StripAll = str
End Function
Function StripDisclaimers(str As String) As String

    Dim r() As String
    Dim x As Integer, y As Integer

    r = Split("Selene Finance LP is a debt collector.|" & _
              "Unless you are in bankruptcy or received a bankruptcy discharge of this debt, Selene Finance is attempting to collect a debt and any information obtained will be used for that purpose.|" & _
              "If you are in bankruptcy or received a bankruptcy discharge of this debt, this communication is not an attempt to collect the debt against you personally, but is notice of a possible enforcement of the lien against the collateral property.|" & _
              "The information in this e-mail message is confidential and may also be attorney-client privileged.|" & _
              "The information is intended only for the use of the individual or entity to whom it is addressed.|" & _
              "If you are not the intended recipient, or the employee or agent responsible for delivering it to the intended recipient, please do not read, copy, use or disclose this communication to others.|" & _
              "If you are not the intended recipient, or the employee or the agent responsible for delivering it to the intended recipient, please do not read, copy, use or disclose this communication to others.|" & _
              "Also, please notify the sender by replying to this message, and then delete it from your system.|" & _
              "For more information about Selene Finance, please visit www.selenefinance.com.|" & _
              "Thank you.|" & _
              "The information is intended only for the  use of the individual or entity to whom it is addressed.|" & _
              "", _
              "|")

    For x = LBound(r) To UBound(r)
        str = Replace(str, r(x), "")
    Next
    
    StripDisclaimers = str

End Function
Function StripMultipleLines(str As String) As String
    Dim cnt As Integer
    
    Do While True
        cnt = 0
        If InStr(str, Chr(13) & Chr(13) & Chr(13)) > 0 Then
            str = Replace(str, Chr(13) & Chr(13) & Chr(13), Chr(13) & Chr(13))
            cnt = 1
        End If
        
        If cnt <> 1 Then: Exit Do

    Loop
    
    StripMultipleLines = str
End Function
Function StripSingleSpaceLines(str As String) As String
    Dim cnt As Integer
    Dim str_r
    Dim x As Long, y As Long, z As Long
    
    Do While True
        cnt = 0
        'eliminate lines with a single space only
        str_r = Split(str, Chr(13)) 'split to array on chr(13) (lf)
        str = "" 'clear out str
        For x = LBound(str_r) To UBound(str_r)
            If str_r(x) = " " Then
                str_r(x) = "" 'replace if line is only a space
                cnt = 1
            End If
        Next
        
        For z = UBound(str_r) To LBound(str_r) Step -1
            If str_r(z) <> "" Then
                Exit For
            End If
        Next
        
        For y = LBound(str_r) To z 'rebuild str
            str = str & str_r(y) & Chr(13)
        Next
        
        If cnt <> 1 Then: Exit Do

        pubCnt = 1
    Loop
    
    StripSingleSpaceLines = str
    
End Function
Function StripSomeThings(str As String) As String
    Dim cnt As Integer
    
    Do While True
        cnt = 0
        'replace nbsp
        If InStr(str, Chr(160)) > 0 Then
            str = Replace(str, Chr(160), Chr(32))
            cnt = 1
        End If
        'replace double spaces
        If InStr(str, "  ") > 0 Then
            str = Replace(str, "  ", " ")
            cnt = 1
        End If
        'replace cr with lf
        If InStr(str, Chr(10)) > 0 Then
            str = Replace(str, Chr(10), Chr(13), vbTextCompare)
            cnt = 1
        End If
        'replace crlf with lf
        If InStr(str, Chr(10) & Chr(13)) > 0 Then
            str = Replace(str, Chr(10) & Chr(13), Chr(13), vbTextCompare)
            cnt = 1
        End If

        If cnt <> 1 Then: Exit Do

        pubCnt = 1
    Loop
    
    StripSomeThings = str
    
End Function


