Attribute VB_Name = "PopUPS"
Sub TempMsg(ByVal item As Object)
Dim Msg As Outlook.MailItem
Dim att As Outlook.Attachment
Dim x
Dim y As Integer
Set Msg = item
Set x = CreateObject("WScript.Shell")
Dim attCounter As Integer
attCounter = 0

For Each att In Msg.Attachments
        
            If InStr(1, att.FileName, "Daily") > 0 And InStr(1, att.FileName, ".xlsm") Then
                
                attCounter = attCounter + 1
                
            End If

Next

If attCounter > 0 Or Msg.SenderEmailAddress = "mubar@nea.com" Then

    y = x.Popup("Do you want to update files come from " & Msg.SenderName & "?", 9, "New Data Mail Received", vbYesNo)
    If y = vbYes Then
        
        Call AutomailStart(Msg)
        
    ElseIf y = vbNo Then
        
        Exit Sub
        
    Else
    
        Call AutomailStart(Msg)
        
        
    End If
    
    'MsgBox "The files come from " & Msg.SenderEmailAddress & " are UPDATED!" _
            & Chr(10) & "Update Time" & Chr(10) & Time, _
            vbOKOnly + vbDefaultButton1, "UPDATE COMPLETED"
    
    Set x = Nothing
    
End If

End Sub



