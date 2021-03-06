Attribute VB_Name = "S_InformationGroup"
Option Private Module


'Callback for InformationButton onAction
Sub Info_Click(control As IRibbonControl)

    a$ = ""
    a$ = a$ & ProgramName$ & " - Information" & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Programmed and Developed from 2016-" & Year(Date) & "." & vbCrLf
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Used under agreed terms." & vbCrLf
    
    a$ = a$ & "Version S" & Format(SharedVersion, "000") & UCase(Left(Replace(Application.Name, "Microsoft ", ""), 1)) & Format(ApplicationVersion, "000") & vbCrLf
    
    a$ = a$ & "" & vbCrLf
    a$ = a$ & "Bug-Reports should be logged via:" & vbCrLf
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



Sub Unavailable(dummy$)

    Result = MsgBox("Function not yet available.", vbCritical, ProgramName$)
    
End Sub

