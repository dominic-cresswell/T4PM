Attribute VB_Name = "RibbonCalls1"


'Callback for InformationButton onAction
Sub Info_Click(control As IRibbonControl)

    a$ = ""
    a$ = a$ & ProgramName$ & " - Information" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Programmed under commission by;" & vbCrLf
    a$ = a$ & "  Dominic Cresswell in 2016." & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Used under agreed terms." & vbCrLf
    
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "E-mail:" & vbCrLf
    a$ = a$ & " dominic.cresswell@ultimateamiga.co.uk" & vbCrLf
    
     MsgBox a$, vbInformation, ProgramName$


End Sub

'Callback for HelpButton onAction
Sub Help_Click(control As IRibbonControl)
    a$ = ""
    a$ = a$ & ProgramName$ & " - Help" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & ".xlam 'add-in' to be installed via developer tab." & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Ideally located at:" & vbCrLf
    a$ = a$ & "C:\Users\*username*\AppData\Roaming\Microsoft\AddIns\" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "" & vbCrLf
    
    MsgBox a$, vbExclamation, ProgramName$
 
 
End Sub

