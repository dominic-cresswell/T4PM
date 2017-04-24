Attribute VB_Name = "S_EmailCode"


Public Mail_from$, Mail_to$(40), Mail_cc$(40), Mail_attach$(40), Mail_savepath$
Public Invite_to$(40)
Public Mail_subject$, Mail_body$
Public FMname$, FMtitle$, FMtele$, FMemail$





Public Sub NewMail()
Dim oOutlook As Object
Dim oMailItem As Object
Dim oRecipient As Object
Dim oNameSpace As Object

' this requires the reference module
'Set oOutlook = New Outlook.Application

'this does not
Set oOutlook = CreateObject("Outlook.Application", "localhost")
'Set oOutlook = CreateObject("Outlook.Application")


Set oNameSpace = oOutlook.GetNamespace("MAPI")
'oNameSpace.Logon , , True
 oNameSpace.Logon "Outlook", , True, False
 
Set oMailItem = oOutlook.CreateItem(0)


' "save to" folder
    If Mail_savepath$ <> "" And FileExists(Mail_savepath$) = True Then
     AddSavePath$ = "<br><br><table bgcolor=#ffffff border=0 cellpadding=0 cellspacing=0><font color=#ffffff face=arial size=" & Chr(34)
     AddSavePath$ = AddSavePath$ & "-6" & Chr(34) & ">SaveToMSF=" & Chr(34) & Mail_savepath$ & Chr(34) & "</font></table>"
    Else
     AddSavePath$ = ""
    End If
    
    

For a = 0 To 40
    If Mail_to$(a) <> "" Then
        Set oRecipient = _
        oMailItem.Recipients.Add(Mail_to$(a))
        oRecipient.Type = 1 '1 = To, use 2 for cc
        'keep repeating these lines with
        'your names, adding to the collection.
    End If
Next a

For a = 0 To 40
    If Mail_cc$(a) <> "" Then
        Set oRecipient = _
        oMailItem.Recipients.Add(Mail_cc$(a))
        oRecipient.Type = 2 '1 = To, use 2 for cc
        'keep repeating these lines with
        'your names, adding to the collection.
    End If
Next a



With oMailItem
    .SentOnBehalfOfName = Mail_from$
    .Subject = Mail_subject$
 '  .BodyFormat = olFormatHTML
     .Display
    .HTMLBody = Mail_body$ & .HTMLBody
    .HTMLBody = .HTMLBody & AddSavePath$
    
    For a = 0 To 20
      If Mail_attach(a) <> "" Then
      .Attachments.Add (Mail_attach$(a)) 'change to your filename
      End If
    Next a
    
 '   Set .SaveSentMessageFolder = fldFac
 
     .Display 'use .Send when all testing done
End With

Mail_from$ = ""
Mail_subject$ = ""
Mail_body$ = ""
For a = 0 To 40: Mail_to$(a) = "": Next a
For a = 0 To 40: Mail_cc$(a) = "": Next a
For a = 0 To 40: Mail_attach$(a) = "": Next a

End Sub

' ===========================================


' ADD SEND ITEMS (ATTACHEMENTS, TO, CC ETC) TO EMAIL

Sub AddressTo(Address$)
If Address$ = "" Or Address$ = Null Then Exit Sub

For a = 0 To 40
    If Mail_to$(a) = "" Or Mail_to$(a) = Address$ Then
    Mail_to$(a) = Address$
    Exit Sub
    End If
Next a

End Sub


Sub InviteTo(Address$)
If Address$ = "" Or Address$ = Null Then Exit Sub

For a = 0 To 40
    If Invite_to$(a) = "" Or Invite_to$(a) = Address$ Then
    Invite_to$(a) = Address$
    Exit Sub
    End If
Next a

End Sub


Sub AddressCC(Address$)
If Address$ = "" Or Address$ = Null Then Exit Sub

For a = 0 To 30
    If Mail_cc$(a) = "" Or Mail_cc$(a) = Address$ Then
    Mail_cc$(a) = Address$
    Exit Sub
    End If
Next a

End Sub


Sub AttachFile(File$)
If File$ = "" Or File$ = Null Then Exit Sub

For a = 0 To 30
    If Mail_attach$(a) = "" Or Mail_attach$(a) = File$ Then
    Mail_attach(a) = File$
    Exit Sub
    End If
Next a

End Sub



' ======================================


Sub MailGreeting()
'    Mail_body$ = Mail_body$ & "<br><font face=arial size=2>"
    Mail_body$ = Mail_body$ & "<br>"
    If Hour(Time) >= 18 Then
      Mail_body$ = Mail_body$ & "Good Evening,"
    ElseIf Hour(Time) >= 12 And Hour(Time) < 18 Then
      Mail_body$ = Mail_body$ & "Good Afternoon,"
    Else
      Mail_body$ = Mail_body$ & "Good Morning,"
    End If
 '   Mail_body$ = Mail_body$ & "</font><br><br>"
    Mail_body$ = Mail_body$ & "<br><br>"

End Sub


'=================
Sub MailSig()
   ' Mail_body$ = Mail_body$ & "<font face=arial size=1>"
     
    a$ = "<br>Kind regards,"
    
    Mail_body$ = Mail_body$ & a$
   ' Mail_body$ = Mail_body$ & "</font>"
End Sub

Sub MailProjectFolder()

On Error GoTo fail:
    Mail_savepath$ = [Form_General].[Folder Path]
    Exit Sub


fail:
    Mail_savepath$ = ""
End Sub



