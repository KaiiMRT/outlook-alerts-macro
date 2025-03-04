Attribute VB_Name = "RoomAlertsMacro"
Option Explicit

Dim WithEvents myInbox As Outlook.Items

Private Sub Application_Startup()
    Dim ns As Outlook.NameSpace
    Dim sharedMailbox As Outlook.Folder
    Dim alertsFolder As Outlook.Folder

    ' Debug message: Macro started successfully
    MsgBox "The macro has started successfully.", vbInformation, "Debug"

    ' Connect to MAPI namespace
    Set ns = Application.GetNamespace("MAPI")
    
    ' Connect to the shared mailbox CGO MGS
    On Error Resume Next
    Set sharedMailbox = ns.Folders.Item("CGO MGS")
    On Error GoTo 0

    ' Check if the shared mailbox exists
    If sharedMailbox Is Nothing Then
        MsgBox "Could not access the shared mailbox CGO MGS.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Access the "Room Alerts" folder within the shared mailbox
    On Error Resume Next
    Set alertsFolder = sharedMailbox.Folders.Item("Inbox").Folders.Item("Room Alerts")
    On Error GoTo 0

    ' Check if the "Room Alerts" folder exists
    If alertsFolder Is Nothing Then
        MsgBox "The 'Room Alerts' folder does not exist in the shared mailbox.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Monitor emails within the "Room Alerts" folder
    Set myInbox = alertsFolder.Items
End Sub

Private Sub myInbox_ItemAdd(ByVal Item As Object)
    ' Debug message: New email detected in the "Room Alerts" folder
    MsgBox "A new email has been detected in the Room Alerts folder.", vbInformation, "Debug"

    ' This code runs when a new email arrives in the "Room Alerts" folder
    If TypeName(Item) = "MailItem" Then
        ' Debug message: The item is a MailItem
        MsgBox "The item is a MailItem.", vbInformation, "Debug"

        ' Check if the email is from "comunicacionesti@mgs.es"
        If Item.SenderEmailAddress = "comunicacionesti@mgs.es" Then
            ' Show a notification
            MsgBox "New email from comunicacionesti@mgs.es in the Room Alerts folder!", vbInformation, "New Email Notification"
        Else
            ' Debug message: The email is not from comunicacionesti@mgs.es
            MsgBox "The email is not from comunicacionesti@mgs.es.", vbInformation, "Debug"
        End If
    Else
        ' Debug message: The item is not a MailItem
        MsgBox "The item is not a MailItem.", vbInformation, "Debug"
    End If
End Sub