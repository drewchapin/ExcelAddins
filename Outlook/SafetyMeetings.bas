Attribute VB_Name = "SafetyMeetings"
Option Explicit

Public Sub AcceptSafetyMeetings()

    ' MAPI Namespace
    Dim MAPI As Outlook.NameSpace
    Set MAPI = Outlook.Application.GetNamespace("MAPI")
    
    ' Inbox folder
    Dim Inbox As Folder
    Set Inbox = MAPI.GetDefaultFolder(olFolderInbox)
    
    ' Loop through each Item in Inbox
    Dim Item As Object
    For Each Item In Inbox.Items
    
        ' Only process MeetingItem's
        If TypeName(Item) = "MeetingItem" Then
        
            ' Get MeetingItem
            Dim Meeting As MeetingItem
            Set Meeting = Item
            
            ' Only process Safety Training meetings
            If InStr(1, Meeting.Subject, "Safety Training", vbTextCompare) > 0 Then
                
                ' Get Appointment
                Dim Appointment As AppointmentItem
                Set Appointment = Meeting.GetAssociatedAppointment(True)
                
                ' Set the category and reminder time
                Appointment.Categories = "Training / Employee Meeting"
                Appointment.ReminderMinutesBeforeStart = 18 * 60
                
                ' Respond
                Dim Response As MeetingItem
                Set Response = Appointment.Respond(olMeetingTentative, True)
                Response.Send
                
                ' Delete MeetingItem
                Meeting.Delete
            
            End If
        
        End If
    
    Next

End Sub
