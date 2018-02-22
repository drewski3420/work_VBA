Option Explicit

Sub MeetingFromEmail()
    
    Dim ol As New Outlook.Application
    Dim olMail As MailItem
    Dim olMeeting As AppointmentItem

    Dim objSel As Object, objSel2 As Object 'Word.Selection
    
    Dim a As Attachment
    Dim fn As String
    On Error Resume Next
    Set olMail = Application.ActiveExplorer.Selection.Item(1)
    If Err.Number <> 0 Then: Exit Sub
    Err.Clear
    On Error GoTo 0
    Set olMeeting = ol.CreateItem(olAppointmentItem)
    
    Set objSel = olMail.GetInspector.WordEditor.Windows(1).Selection
    objSel.wholestory
    objSel.Copy

    Set objSel2 = olMeeting.GetInspector.WordEditor.Windows(1).Selection
    objSel2 = "From: " & olMail.Sender & vbCrLf & _
                   "Sent: " & olMail.SentOn & vbCrLf & _
                   "To: " & olMail.To & vbCrLf & _
                   IIf(olMail.CC = "", "", "CC: " & olMail.CC) & vbCrLf & vbCrLf & _
                   StripAll(objSel.Text)  '.Text

    With olMeeting
        .Subject = "Created from Email: " & olMail.ConversationTopic
        .start = Date + 1 + (((Int(Timer / 900) + 1) * 900) / 86400)
        .Duration = 0
        .ReminderMinutesBeforeStart = 0
        .Display
        For Each a In olMail.Attachments
            fn = Environ$("tmp") & "\" & a.fileName
            a.SaveAsFile fn
            olMeeting.Attachments.Add fn
        Next
        .Save
        '.Close (olSave)
    End With
End Sub
