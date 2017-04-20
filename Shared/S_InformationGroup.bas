Attribute VB_Name = "S_InformationGroup"


'Callback for InformationButton onAction
Sub Info_Click(control As IRibbonControl)

    a$ = ""
    a$ = a$ & ProgramName$ & " - Information" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Programmed and Developed by;" & vbCrLf
    a$ = a$ & "  Dominic Cresswell in 2016-" & Year(Date) & "." & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Used under agreed terms." & vbCrLf
    
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "E-mail:" & vbCrLf
    a$ = a$ & " dominic.cresswell@gen2.co.uk" & vbCrLf
    
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Bug-Reports:" & vbCrLf
    a$ = a$ & " https://github.com/dominic-cresswell/T4PM" & vbCrLf

    a$ = a$ & "" & vbCrLf
    a$ = a$ & "User Data Path:" & vbCrLf
    a$ = a$ & ProgramPath$ & vbCrLf
    
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

