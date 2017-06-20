Attribute VB_Name = "O_RibbonButtonCode"

' outlook intercepts on button presses
Sub Store_Emails()
    Call gsubExportMail
End Sub


Sub Tag_Emails()

Call Unavailable("")

End Sub


Sub PickProject_Click()

Call Unavailable("")

End Sub


Sub Email_Click()

Call Unavailable("")

End Sub


Sub Folder_Click()

Call Unavailable("")

End Sub


Sub Unavailable(dummy$)

    Result = MsgBox("Function currently not yet implemented.", vbCritical, ProgramName$)
    

End Sub
