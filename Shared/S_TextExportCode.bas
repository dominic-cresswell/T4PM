Attribute VB_Name = "S_TextExportCode"


Function ClearSpecialCharacters(incomeString$)

inString$ = incomeString$

On Error Resume Next

    inString$ = Replace(inString$, " ", "")
    inString$ = Replace(inString$, "/", "")
    inString$ = Replace(inString$, "\", "")
    inString$ = Replace(inString$, "_", "")
    inString$ = Replace(inString$, ":", "")
    inString$ = Replace(inString$, "&", "")
    inString$ = Replace(inString$, "+", "")
    inString$ = Replace(inString$, "!", "")
    inString$ = Replace(inString$, "@", "")
    inString$ = Replace(inString$, "#", "")
    inString$ = Replace(inString$, "(", "")
    inString$ = Replace(inString$, ")", "")
    inString$ = Replace(inString$, "{", "")
    inString$ = Replace(inString$, "}", "")
    inString$ = Replace(inString$, "[", "")
    inString$ = Replace(inString$, "]", "")
    inString$ = Replace(inString$, ",", "")
    inString$ = Replace(inString$, ".", "")
    inString$ = Replace(inString$, ";", "")
    inString$ = Replace(inString$, "£", "")
    inString$ = Replace(inString$, "$", "")
    inString$ = Replace(inString$, "*", "")
    inString$ = Replace(inString$, "^", "")
    inString$ = Replace(inString$, "%", "")
    inString$ = Replace(inString$, "~", "")
    inString$ = Replace(inString$, "`", "")
    inString$ = Replace(inString$, "|", "")
    inString$ = Replace(inString$, "±", "")
    inString$ = Replace(inString$, "§", "")
    
    
' ====
  ClearSpecialCharacters = inString$
    
    
End Function




'======

Public Sub MakeVcardV3()
Dim PP$
On Error Resume Next

If CCName$ = "" And CCOrg$ = "" Then GoTo leave:

CCAdd1$ = Replace(CCAdd1$, ";", ",")
CCAdd2$ = Replace(CCAdd2$, ";", ",")
CCAdd3$ = Replace(CCAdd3$, ";", ",")
CCAdd4$ = Replace(CCAdd4$, ";", ",")
    CCPC$ = Replace(CCPC$, ";", ",")

CCNote$ = Replace(CCNote$, vbCr, "\n")
CCNote$ = Replace(CCNote$, vbLf, "")

PP$ = Chr(5)

Dim vCardData As String

    vCardData = ""
    vCardData = vCardData & "BEGIN:VCARD" & PP$
    vCardData = vCardData & "VERSION:3.0" & PP$
    vCardData = vCardData & "PRODID:-//DJC//Dominic_Cresswell//EN" & PP$
    If CCOrg$ <> "" Then vCardData = vCardData & "ORG;type=pref:" & CCOrg$ & PP$
  '   If CCName$ = "" And CCOrg$ <> "" Then vCardData = vCardData & "FN;type=pref:" & CCName$ & PP$
    If CCName$ <> "" Then vCardData = vCardData & "N:" & CCName$ & PP$
    
    If CCAdd1$ <> "" Then vCardData = vCardData & "item1.ADR;type=WORK;type=pref:;;" & CCAdd1$ & ";" & CCAdd2$ & ";" & CCAdd3$ & ";" & CCAdd4$ & ";" & CCPC$ & PP$
    If CCAdd1$ <> "" Then vCardData = vCardData & "LABEL;WORK;ENCODING=QUOTED-PRINTABLE:" & CCAdd1$ & "=0D=0A" & CCAdd2$ & "=0D=0A" & CCAdd3$ & "=0D=0A" & CCPC$ & PP$
    vCardData = vCardData & "item1.X-ABADR:uk" & PP$
    
    If CCTel$ <> "" Then vCardData = vCardData & "TEL;type=WORK;type=pref:" & CCTel$ & PP$
    If CCFax$ <> "" Then vCardData = vCardData & "TEL;type=FAX;type=pref:" & CCFax$ & PP$
    If CCURL$ <> "" Then vCardData = vCardData & "URL;type=WORK;type=pref:" & CCURL$ & PP$
    If CCEmail$ <> "" Then vCardData = vCardData & "EMAIL;type=INTERNET;type=pref:" & CCEmail$ & PP$

    
     If CCPos$ <> "" Then vCardData = vCardData & "TITLE:" & CCPos$ & PP$
  
    If CCNote$ <> "" Then vCardData = vCardData & "NOTE:" & CCNote$ & PP$
    If CCName$ = "" And CCOrg$ <> "" Then vCardData = vCardData & "X-ABShowAs:COMPANY" & "" & PP$
     vCardData = vCardData & "END: VCARD" & PP$
    
  
    If CCName$ <> "" Then CCFName$ = CCName$
    If CCOrg$ <> "" And CCFName$ = "" Then CCFName$ = CCOrg$

    CCFName$ = Replace(CCFName$, "/", "-")
    CCFName$ = Replace(CCFName$, "\", "-")
    
    
    
    Call MakeTextFile(vCardData, CCFPath$ & CCFName$ & ".vcf")



leave:
      CCRef$ = ""
      CCOrg$ = ""
     CCName$ = ""
     CCAdd1$ = ""
     CCAdd2$ = ""
     CCAdd3$ = ""
     CCAdd4$ = ""
       CCPC$ = ""
      CCPos$ = ""

      CCTel$ = ""
      CCMob$ = ""
      CCFax$ = ""
    CCEmail$ = ""
     CCNote$ = ""
      CCURL$ = ""
      
 '   CCFName$ = ""
 '   CCFPath$ = ""
End Sub



Sub MakeTextFile(MyData$, SaveFile$)
'On Error Resume Next

If Len(MyData$) < 1 Then Exit Sub

   Dim fs, a As Variant
   Dim NewCount, CountS As Long
   Dim GetValue$, newLine$, PP$
   
   
   PP$ = Chr(5)
   
   Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(SaveFile$, True)

CountS = Len(MyData$)


For NewCount = 1 To CountS

    GetValue$ = Mid(MyData$, NewCount, 1)

     If GetValue$ <> PP$ And NewCount < CountS Then
            newLine$ = newLine$ + GetValue$
        ElseIf NewCount = CountS Then
            newLine$ = newLine$ + GetValue$
            a.WriteLine newLine$
                        newLine$ = ""
        
        Else
          '  Debug.Print newLine$
            a.WriteLine newLine$
                        newLine$ = ""
        End If
    
Next NewCount

a.Close
End Sub

Function ReadTextFile(inFile$) As String

If FileExists(inFile$) = False Then Exit Function

Dim strFilename As String
strFilename = inFile$
Dim strFileContent As String
Dim iFile As Long: iFile = FreeFile
Open strFilename For Input As #iFile
strFileContent = Input(LOF(iFile), iFile)
ReadTextFile = strFileContent
Close #iFile


End Function
