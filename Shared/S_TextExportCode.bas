Attribute VB_Name = "S_TextExportCode"
Option Private Module


Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias _
"GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
(ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
(ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'~~> Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

Private hHook As Long

Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
    Dim RetVal
    Dim strClassName As String, lngBuffer As Long

    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If

    strClassName = String$(256, " ")
    lngBuffer = 255
    
    If lngCode = HCBT_ACTIVATE Then
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        '~~> Class name of the Inputbox
        If Left$(strClassName, RetVal) = "#32770" Then
            '~~> This changes the edit control so that it display the password character *.
            '~~> You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
    End If

    '~~> This line will ensure that any other hooks that may be in place are
    '~~> called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam

End Function

Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
Optional YPos, Optional HelpFile, Optional Context) As String
    Dim lngModHwnd As Long, lngThreadID As Long
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
    UnhookWindowsHookEx hHook
End Function





' some special text functions

Function CapText(inText$, amount As Long) As String

    CapText = inText$
    If Len(CapText) > amount Then CapText = Left(CapText, amount - 3) & "..."

End Function



Function AddSlash(inString$) As String

    AddSlash = inString$
         If Right(AddSlash, 1) <> ":" And AddSlash <> "" And Right(AddSlash, 1) <> "\" Then
         AddSlash = AddSlash & "\"
        End If

   
End Function



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




'=================================  Export Contact Card

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


'=================================  Export Plain Text File


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

'=================================  Import Plain text File


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
