Dim gEntryId        As String


Function GetCurrentItem() As Object
    Dim objApp      As Outlook.Application
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = _
                objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = _
                objApp.ActiveInspector.CurrentItem
        Case Else
    End Select
End Function


Public Sub MarkMailForMeeting()
    
    Dim myOlApp     As Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim mySel       As Outlook.Selection
    Dim myItem      As Outlook.MailItem
    
    Set mySel = Application.ActiveExplorer.Selection
    Set myItem = mySel.Item(1)
    
    gEntryId = myItem.EntryID
End Sub

Public Sub AttachItem()
    'Attach MailItem from gEntryId into the currently selected Appointment
    Dim olApp       As Outlook.Application
    Dim olApt       As Outlook.AppointmentItem
    Dim olMail      As Outlook.MailItem
    Dim olMailCopy  As Outlook.MailItem
    Dim myCopiedMessage As Outlook.MailItem
    Dim olNS        As Outlook.NameSpace
    
    Set olApp = New Outlook.Application
    '   Set olApt = olApp.CreateItem(olAppointmentItem)
    
    Set olNS = olApp.GetNamespace("MAPI")
    
    olNS.Logon
    
    Set olMail = olNS.GetItemFromID(gEntryId)
    Set olMailCopy = olMail.Copy
    
    Dim olSel       As Outlook.Selection
    Set olSel = Application.ActiveExplorer.Selection
    
    olSel.Item(1).Attachments.Add olMailCopy, olByValue
    '    olItem.Attachments.Add olMailCopy, olByValue
    olSel.Item(1).Save
    olMailCopy.Delete
    
End Sub

Public Sub ToDo_WaitingFor()
    tFolder = "Waiting For"
    CreateTasks (tFolder)
End Sub
Public Sub ToDo_Next()
    tFolder = "Next"
    CreateTasks (tFolder)
End Sub
Public Sub ToDo_SmallP()
    tFolder = "Small"
    CreateTasks (tFolder)
End Sub
Public Sub ToDo_LargeP()
    tFolder = "Large"
    CreateTasks (tFolder)
End Sub
Public Sub ToDo_Later()
    tFolder = "Later"
    CreateTasks (tFolder)
End Sub

Private Sub CreateTasks(tFolder As String)
    
    Dim Ns          As Outlook.NameSpace
    Dim olTask      As Outlook.TaskItem
    Dim Item        As Outlook.MailItem
    Dim taskFolders As Outlook.Folders
    Dim taskFolder  As Outlook.Folder
    
    Set Ns = Application.GetNamespace("MAPI")
    
    Set Item = Application.ActiveExplorer.Selection.Item(1)
    
    Set taskFolders = Ns.GetDefaultFolder(olFolderTasks).Folders
    For Each Folder In taskFolders
        If InStr(1, Folder.Name, tFolder, 1) > 0 Then
            Set taskFolder = Folder
            Exit For
        End If
    Next
    
    Item.Move taskFolder.Parent.Parent.Folders("Archive")
    
    Set olTask = taskFolder.Items.Add(olTaskItem)
    With olTask
        .Subject = Item.Subject
        .Attachments.Add Item
        If tFolder = "Next" Then
            .ReminderSet = True
            .ReminderTime = Now + TimeSerial(0, 5, 0)
            
        End If
        
        
        .Body = Item.Body
        .RTFBody = Item.RTFBody
        .Save
    End With
    
    Set Ns = Nothing
End Sub



Public Sub SaveAttachment()
    Dim objItem     As Outlook.MailItem
    Dim objAttach   As Outlook.Attachment
    Dim dateFormat  As String
    dateFormat = Format(Now, "yyyy-mm-dd H-mm")
    
    Set objItem = GetCurrentItem()
    Const strAttachmentPath = "C:\Users\juergen.richert\OneDrive - BioNTech SE\Attachments\"
    
    For i = 1 To objItem.Attachments.Count
        Set objAttach = objItem.Attachments(i)
        
        objAttach.SaveAsFile strAttachmentPath & dateFormat & " " & objAttach.FileName
    Next
    
    Set objItem = Nothing
    Set objMail = Nothing
End Sub





Public Sub BulkDeleteAppointments()
    Dim oAppt As Object
    Dim itemsToDelete As Object
    Dim cancelMsg As String

    ' did user select date range or individual items?
    
    If Application.ActiveExplorer.Selection.Count > 0 Then
        Set itemsToDelete = Application.ActiveExplorer.Selection
    ElseIf (datRange.startDate <> datRange.datNull) And (datRange.endDate <> datRange.datNull) Then
        MsgBox "Nothing selected.", vbOKOnly, "Bulk delete"
        Exit Sub
    End If
    ' only proceed if items are selected by now, either automatically or manually
    If itemsToDelete.Count > 0 Then
        ' What shall we send as message?
        cancelMsg = InputBox(Prompt:="Your cancel message please. There will be no confirmation.", _
              Title:="ENTER YOUR MESSAGE", Default:="I will be on vacation.")
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
    If (strMessageClass = "IPM.Appointment") Then       ' Only operate on Calendar Entry.
        Set oAppointItem = oItem
        If oAppointItem.Organizer = Outlook.Session.CurrentUser Then  ' If this is my own meeting
            oAppointItem.MeetingStatus = olMeetingCanceled
            oAppointItem.Body = cancelMsg
            oAppointItem.Save
            oAppointItem.Send
        Else                                            ' If I was invited to this meeting
            Set myMtg = oAppointItem.Respond(olMeetingDeclined, True, False)
            If Not myMtg Is Nothing Then
                myMtg.Body = cancelMsg
                myMtg.Send
            End If
        End If
    End If
End Sub

Public Sub BulkDeleteAppointments_noMessage()
    Dim oAppt As Object
    Dim itemsToDelete As Object
    Dim cancelMsg As String

    ' did user select date range or individual items?
    
    If Application.ActiveExplorer.Selection.Count > 0 Then
        Set itemsToDelete = Application.ActiveExplorer.Selection
    ElseIf (datRange.startDate <> datRange.datNull) And (datRange.endDate <> datRange.datNull) Then
        MsgBox "Nothing selected.", vbOKOnly, "Bulk delete"
        Exit Sub
    End If
    ' only proceed if items are selected by now, either automatically or manually
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
    If (strMessageClass = "IPM.Appointment") Then       ' Only operate on Calendar Entry.
        Set oAppointItem = oItem
        If oAppointItem.Organizer = Outlook.Session.CurrentUser Then  ' If this is my own meeting
            ' Do nothing
        Else                                            ' If I was invited to this meeting
            Set myMtg = oAppointItem.Respond(olMeetingDeclined, True, False)
            If Not myMtg Is Nothing Then
                myMtg.Save
            End If
        End If
    End If
End Sub
