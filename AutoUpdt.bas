Attribute VB_Name = "autoUpdt"
Function programStart() As Workbook
        Dim exapp As Excel.Application
        Dim ExWbk As Workbook
        Set exapp = New Excel.Application
        Set ExWbk = exapp.Workbooks.Open("J:\My Drive\Gkr\program.xlsm", _
        UpdateLinks:=0)
        exapp.Visible = True
        'ExWbk.Application.Run ("Copy_Materials.Material_List_Combination")
    Set programStart = ExWbk
End Function

Sub AutomailStart(ByVal item As Object)
    Dim SenderAdress As String
    Dim Msg As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim attlist As Scripting.Dictionary
    Set attlist = New Dictionary
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim flnm As String
    Dim rowno As Integer
    Dim j As Long
    Dim i As Long
    On Error Resume Next
    Set Msg = item
    SenderAdress = Msg.SenderEmailAddress
    
    
    
    If SenderAdress = "rahim.turk@ne.com" Or SenderAdress = "Nick.fury@n.com" Or SenderAdress = "med.zek@nea.com" Or SenderAdress = "ian.afl@ne.com" Or SenderAdress = "meet.kaas@ma.com" Or SenderAdress = "amed.als@nea.com" Then
    
         '//////////////*******************GOKERRR*******************\\\\\\\\\\\\\\\
        
         i = 0
        
        For Each att In Msg.Attachments
        
            If InStr(1, att.FileName, "Daily") > 0 Then
                flnm = SenderAdress & "  " & Date - 1 & "  " & att.FileName
                attlist.Add Key:=i, item:=flnm
                att.SaveAsFile ("J:\My Drive\Gkr\Reports\Daily Reports\" & flnm)
                i = i + 1
                
            End If

                        
        Next
        Set wb = programStart()
        Set ws = wb.Worksheets("Daily Report Update")
        ws.Range("B1") = flnm
        
        
       
        wb.Application.Run ("Daily_QTY_update")
        
        
        
        
        wb.Close True
        
        '//////////////*******************TUFAIL*******************\\\\\\\\\\\\\\\
        
        
        
        
        '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
          
       
        
        '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
        'muhammad.babar@nesma.com
        
        '//////////////*******************GOKER*******************\\\\\\\\\\\\\\\
        
        
        
    ElseIf SenderAdress = "mambar@nma.com" Then
        
        '//////////////*******************TUFAIL*******************\\\\\\\\\\\\\\\
        
         i = 0
        
        For Each att In Msg.Attachments
        
            If IsInArray(att.FileName, Array("10-40 MM(NEW).xlsb", "0-5 MM= (ALJABRIA)=(BIYOUT ALJEWAR)(NEW).xlsb", _
            "AL DOSARI,BELEDIYE(NEW).xlsb", "AL GHARBI ZONE 5(NEW).xlsb", "KAAR AND MASAR(NEW).xlsb")) Then
                flnm = att.FileName
                attlist.Add Key:=i, item:=flnm
                att.SaveAsFile ("J:\My Drive\Gkr\Tufail\" & att.FileName)
                i = i + 1
                
            End If

                        
        Next
        Set wb = programStart()
        Set ws = wb.Worksheets("Data")
        
        ws.Range("B4:B20").Delete
        ws.Cells(1, 5) = ""
        
        rowno = 4
        For j = 0 To attlist.Count - 1
            If attlist(j) = "KAAR AND MASAR(NEW).xlsb" Then
                
                ws.Cells(1, 5).Value = attlist(j)
            Else
                
                ws.Cells(rowno, 2).Value = attlist(j)
            End If
           rowno = rowno + 1
           
        Next j
        wb.Application.Run ("Edit_Files.Edit_Materials")
        wb.Application.Run ("Copy_Materials.Material_List_Combination")
        
        ws.Cells(1, 3) = "LAST UPDATE"
        ws.Cells(2, 3) = Date & Chr(10) & Time
        
        wb.Close True
        
        '//////////////*******************TUFAIL*******************\\\\\\\\\\\\\\\
        
        
        
        
        '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
          
         i = 0
        
        For Each att In Msg.Attachments
        
            If IsInArray(att.FileName, Array("EQUIPMENT-HOUR MY COVER(NEW).xlsb")) Then
                flnm = att.FileName
                attlist.Add Key:=i, item:=flnm
                att.SaveAsFile ("J:\My Drive\Gkr\M. BABAR\" & att.FileName)
                i = i + 1
                
            End If
                        
        Next
        
        
        Set wb = programStart()
        Set ws = wb.Worksheets("Our EQ Timesheets")
        
        
        wb.Application.Run ("Our_EQ_List.MBabarList")
        
        
        ws.Cells(1, 3) = "LAST UPDATE"
        ws.Cells(2, 3) = Date & Chr(10) & Time
        
        
        '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
        
        
    
    'ElseIf SenderAdress = "ammbr@nma.com" Then
    
    
    '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
          
'         i = 0
'
'        For Each att In Msg.Attachments
'
'            If IsInArray(att.FileName, Array("EQUIPMENT-HOUR MY COVER.xlsx")) Then
'                flnm = att.FileName
'                attlist.Add Key:=i, item:=flnm
'                att.SaveAsFile ("J:\My Drive\Gkr\M. BABAR\" & att.FileName)
'                i = i + 1
'
'            End If
'
'        Next
'
'
'        Set wb = programStart()
'        Set ws = wb.Worksheets("Our EQ Timesheets")
'
'
'        wb.Application.Run ("Our_EQ_List.MBabarList")
'
'
'        ws.Cells(1, 3) = "LAST UPDATE"
'        ws.Cells(2, 3) = Date & Chr(10) & Time
'
        
        '//////////////*******************BABAR*******************\\\\\\\\\\\\\\\
    
    
        
    End If
    
    wb.Close (True)

End Sub


Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
