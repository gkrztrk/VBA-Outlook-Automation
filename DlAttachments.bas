Attribute VB_Name = "DlAttachments"
Sub dlAttachments()

Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Set olApp = Outlook.Application
Dim olFolder As Outlook.MAPIFolder
Set objNS = olApp.GetNamespace("MAPI")
Dim Msg As Outlook.MailItem
Dim att As Outlook.Attachment
Dim attlist As Scripting.Dictionary

    Set attlist = New Dictionary

  ' default local Inbox
  Set olFolder = objNS.GetDefaultFolder(olFolderInbox).Folders("Timekeeper")
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Folders("Timekeeper").Items

 If TypeName(Items) = "MailItem" Then
    Set Msg = Items
 End If


    For Each item In olFolder.Items
        
        If TypeName(item) = "MailItem" Then
            
            Set Msg = item

       
        
        For Each att In Msg.Attachments
        
        
            If InStr(att.FileName, "Zone Wise") > 0 Then
                oPath = "J:\My Drive\Gkr\Data Source\employers\" & CStr(Msg.ReceivedTime) & ".xlsx"
                att.SaveAsFile oPath
            End If
        Next
        
        End If
    Next

End Sub
