Public Sub BulkDeleteAppointments()
    Dim oAppt As Object
    Dim itemsToDelete As Object
    Dim cancelMsg As String
    
    
    If Application.ActiveExplorer.Selection.Count > 0 Then
        Set itemsToDelete = Application.ActiveExplorer.Selection
    Else
        Exit Sub
    End If
    
    
    If itemsToDelete.Count > 0 Then
        cancelMsg = InputBox(Prompt
         = "Your cancel message please. There will be no confirmation.", _
        Title
         = "Response", Default
         = "I will be on vacation.")
        If (cancelMsg <> "") Then
            For Each oAppt In itemsToDelete
                DeleteItemWithDefaultMessage oAppt, cancelMsg
            Next oAppt
        End If
    End If
End Sub

Private Sub DeleteItemWithDefaultMessage(oItem, cancelMsg)
    Dim strMessageClass As String
    Dim oAppointItem As Outlook.AppointmentItem
    Dim myMtg As Outlook.MeetingItem
    strMessageClass = oItem.MessageClass
    If (strMessageClass = "IPM.Appointment") Then
        Set oAppointItem = oItem
        If oAppointItem.Organizer = Outlook.Session.CurrentUser Then  ' If this is my own meeting
        oAppointItem.MeetingStatus = olMeetingCanceled
        oAppointItem.Body = cancelMsg
        oAppointItem.Save
        oAppointItem.Send
    Else
        Set myMtg = oAppointItem.Respond(olMeetingDeclined, True, False)
        If Not myMtg Is Nothing Then
            myMtg.Body = cancelMsg
            myMtg.Send
        End If
    End If
End If
End Sub

' Delete without any message
Public Sub BulkDeleteAppointments_noMessage()
Dim oAppt As Object
Dim itemsToDelete As Object
Dim cancelMsg As String



If Application.ActiveExplorer.Selection.Count > 0 Then
    Set itemsToDelete = Application.ActiveExplorer.Selection
Else
    Exit Sub
End If

If itemsToDelete.Count > 0 Then
    For Each oAppt In itemsToDelete
        DeleteItemWithoutMessage oAppt
    Next oAppt
    
End If
End Sub

Private Sub DeleteItemWithoutMessage(oItem)
Dim strMessageClass As String
Dim oAppointItem As Outlook.AppointmentItem
Dim myMtg As Outlook.MeetingItem

strMessageClass = oItem.MessageClass

If (strMessageClass = "IPM.Appointment") Then
    Set oAppointItem = oItem
    If oAppointItem.Organizer = Outlook.Session.CurrentUser Then
        ' Do nothing
    Else
        Set myMtg = oAppointItem.Respond(olMeetingDeclined, True, False)
        If Not myMtg Is Nothing Then
            myMtg.Save
        End If
    End If
End If
End Sub
