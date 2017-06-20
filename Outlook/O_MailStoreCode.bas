Attribute VB_Name = "O_MailStoreCode"



Public Sub gsubExportMail()
'PURPOSE: Exports selected emails to a selected Explorer folder
'Peter Riley
'19 Nov 2008 - modified to work around non-mail items in selected list

Dim selA As Outlook.Selection
Dim itemA As Object
Dim email As Outlook.MailItem
Dim iEmails As Integer
Dim iNonEmails As Integer

Dim sSubject As String
Dim sSender As String
Dim dtReceived As Date

Dim sPath As String
Dim sFileName As String
Dim foundPath As String

Dim sMsg As String
Dim flagType As String

'following code amended from original for KCC Outlook 2010 - create xlApp as opposed to With Excel.Application
Dim xlApp As Object

Dim i As Integer
Dim sBan As String    'banned chars in filenames
    sBan = "\:/*?<>|" & Chr(34) & vbLf & vbCr         '34 is double quotes (")

'PRINCIPLES:
'Outlook.ActiveExplorer.Selection gives the emails selected
'File dialog allows target folder to be selected

'START
'Select email
   Set selA = Outlook.ActiveExplorer.Selection
      If selA.Count = 0 Then
         MsgBox "You must select some emails to be exported."
         Exit Sub
   End If
  

    Result = MsgBox("Please select a folder for manual storage of e-mails, " & vbCrLf & " or click 'cancel' to skip and store T4PM e-mails only.", vbInformation + vbOKCancel, ProgramName$)
    
    
    ' no to manual selection!
    If Result = vbCancel Then
        sPath = "<?>"
        
    Else

'ooh that's a cheeky bit of code!

'Select target folder by dialog box (this is only available thro Excel !!)
'Was: With Excel.Application.FileDialog(msoFileDialogFolderPicker)
    Set xlApp = CreateObject("excel.application")
    With xlApp.FileDialog(msoFileDialogFolderPicker)
         
     'Set button caption
     .ButtonName = "Select"
   
     'Allow several files to be selected
  'Was: .AllowMultiSelect = False ' should this be True?
     .AllowMultiSelect = True
        
     'Use 'Show' to display the File Picker dialog box and return the selected folder
     
      If .Show = -1 Then 'The user pressed the action button.
         sPath = .SelectedItems(1) & "\"      'selected folder with \ added
      
      Else     'user cancelled
         sPath = "<?>"
      End If
   
   End With
   
'finished with dialog box
'Was: Excel.Application.Quit
    xlApp.Quit
    
    
    ' did not want to select a manual path
    End If
    

      If sPath <> "<?>" Then
          sPath = AddSubPath(sPath)
      End If




'Open email folder & process emails (Folders are collections in collections)
  iEmails = 0
  iNonEmails = 0
  skipMail = 0
  For Each itemA In selA
  
  
  If itemA.Class = olMail Then   'check that item selected is an email (not calendar appoitnment etc)
 
                    'Count and set email
                  '  iEmails = iEmails + 1
                    Set email = itemA
                    Debug.Print iEmails, email.Subject, , email.ReceivedTime
                    
                    'Get heading & date
                    sSubject = Trim(email.Subject)    ' [email] gives the email's own title
                    
                    If Len(sSubject) > 35 Then sSubject = Left(sSubject, 30) + "[...]"
                    dtReceived = email.ReceivedTime
                    sSender = Trim(email.SenderName)
                        
                    'Modify heading
                        'Mark if Untitled
                        If sSubject = "" Then
                          sSubject = "(Untitled)"
                        End If
                        
                        If sSender = "" Then
                            sSender = "(No Sender)"
                        End If
                        
                        'Remove banned chars from heading      \ / : * ? " < > |  Cr Lf
                        For i = 1 To Len(sBan)
                          sSubject = Replace(sSubject, Mid(sBan, i, 1), " ")    'replace any banned chars with space
                          sSender = Replace(sSender, Mid(sBan, i, 1), " ")
                        Next i
                    
                    'Create filename
                    ' was 100 but throwing errors with legal's filenames, reduced to 75 then 50, still some issues...
                '       sFileName = Format(dtReceived, "yymmdd") & " " & Format(dtReceived, "hhmmss") & " " & email.SenderName & " " & Left(sSubject, 100) & ".msg"
                    sFileName = Format(dtReceived, "yymmdd") & " " & Format(dtReceived, "hhmmss") & " " & email.SenderName & " " & Left(sSubject, 100) & ".msg"
                        
        ' see if we have a path
                    
                foundPath = ""
                
         ' store all e-mails with "SaveToFolder="

                StartPoint = InStr(vbTextCompare, itemA.Body, "SaveToFolder=")
                If StartPoint > 0 Then
                
                    foundPath = Right(itemA.Body, Len(itemA.Body) - StartPoint + 1)
                    foundPath = Replace(foundPath, "SaveToFolder=" + Chr(34), "")
                    
                    EndPoint = InStr(vbTextCompare, foundPath, Chr(34))
                    foundPath = Left(foundPath, EndPoint - 1)

                    If Right(foundPath, 1) <> "\" Then foundPath = foundPath & "\"


                        If FileExists(foundPath) = False Then
                            foundPath = ""
                        Else
                            flagType = "T4PM Exported"
                        End If
                        
                    Else
                    
                    flagType = "Exported"
                    foundPath = ""
                End If
                

        Debug.Print sFileName
        
        saveDone = False
        If foundPath <> "" Then
            foundPath = AddSubPath(CStr(foundPath))
            email.SaveAs foundPath & sFileName, olMSG
            saveDone = True
            
        ElseIf sPath <> "<?>" Then
             email.SaveAs sPath & sFileName, olMSG
             saveDone = True
        Else
        
        End If


                
   If saveDone = True Then
    'Set "Exported" category (or any of your choice)
        With email
            .Categories = flagType
            .Save   ' if you don't save only first selection is categorised
        End With
        iEmails = iEmails + 1
    Else
        skipMail = skipMail + 1
    End If
    

   Else  'select item not an email (class <> olMail)
      iNonEmails = iNonEmails + 1
      
   End If  'object is an email
  
  Next itemA
  
  'Summary message
   sMsg = iEmails & " emails have been copied to selected path(s)" & vbCrLf ' & sPath
   If skipMail > 0 Then sMsg = sMsg & vbCrLf & vbCrLf & skipMail & " emails were skipped (no path selected)."
   If iNonEmails > 0 Then sMsg = sMsg & vbCrLf & vbCrLf & iNonEmails & " other items not copied."
   
   MsgBox sMsg, vbInformation, "Email Export"

Exit Sub

MailStore:

 

Return



End Sub

Function AddSubPath(inPath$) As String
    
    AddSubPath = inPath$

    If DirExists(AddSubPath + "14 Correspondence\") Then AddSubPath = AddSubPath + "14 Correspondence\"
    If DirExists(AddSubPath + "_. Emails\") Then AddSubPath = AddSubPath + "_. Emails\"
    If DirExists(AddSubPath + "E-Mail\") Then AddSubPath = AddSubPath + "E-Mail\"
    If DirExists(AddSubPath + "E-Mails\") Then AddSubPath = AddSubPath + "E-Mails\"
    If DirExists(AddSubPath + "EMails\") Then AddSubPath = AddSubPath + "EMails\"
    If DirExists(AddSubPath + "EMail\") Then AddSubPath = AddSubPath + "EMail\"

    


End Function

