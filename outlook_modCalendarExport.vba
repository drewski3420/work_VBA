Option Explicit

Public Sub ExportEntireCalendar()
    Dim oNamespace As NameSpace
    Dim oFolder As Folder
    Dim oCalendarSharing As CalendarSharing
    
    ' Get a reference to the Calendar default folder
    Set oNamespace = Application.GetNamespace("MAPI")
    Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' Get a CalendarSharing object for the Calendar default folder.
    Set oCalendarSharing = oFolder.GetCalendarExporter
    ' Set the CalendarSharing object to export the contents of
    ' the entire Calendar folder, including attachments and
    ' private items, in full detail.
    
    With oCalendarSharing
        .CalendarDetail = olFullDetails
        .IncludeWholeCalendar = True
        .IncludeAttachments = True
        .IncludePrivateDetails = True
        .RestrictToWorkingHours = False
    End With
    
    ' Export calendar to an iCalendar calendar (.ics) file.
    oCalendarSharing.SaveAsICal Environ("temp") & "\DrewCalendar.ics"

End Sub
