Attribute VB_Name = "E_RibbonButtonCode"
Option Private Module

Public Const ApplicationVersion = 3


Sub ExcelRibbonLoad()

    RememberProject = False
    
' ~~~~~~~~~~~~~~~~~~
        If S_UserConfigCode.GetConfigSetting("RememberLastProject") <> "" _
           And Replace(LCase(S_UserConfigCode.GetConfigSetting("RememberLastProject")), vbCr, "") = "true" Then
            RememberProject = True
            ' now load the project name from data file
            

            ' now load the project itself
       Else
            RememberProject = False
    
       End If
    
             
End Sub



'Callback for NewProject onAction
Sub NewProject_Click(control As IRibbonControl)

    
    ' check for the four main requirements
    Call PullWriteDataFromWorksheets("")
    
    FailText$ = ""
    
        Test1$ = GetTempData_WriteBuffer("Site Name")
     If Test1$ = "" Then FailText$ = FailText$ & "Site Name details not known." & vbCrLf
        
        Test2$ = GetTempData_WriteBuffer("Project Description")
     If Test2$ = "" Then FailText$ = FailText$ & "Project Description details not known." & vbCrLf
    
        Test3$ = GetTempData_WriteBuffer("Project Manager")
     If Test3$ = "" Then FailText$ = FailText$ & "Project Manager details not known." & vbCrLf
    
        Test4$ = GetTempData_WriteBuffer("Project Reference")
     If Test4$ = "" Then FailText$ = FailText$ & "Project Reference details not known." & vbCrLf


    If FailText$ <> "" Then
        FailText$ = FailText$ & vbCrLf & "Cannot create New Data Store without base information."
        Result = MsgBox(FailText$, vbCritical, ProgramName$)
        Exit Sub
    End If
    
    ' working folder
    
    
    
    
    ' otherwise, we are good to go!
        
        WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))
        
        If WorkingPath$ = "" Or DirExists(WorkingPath$) = False Then
            Result = MsgBox("Working Folder Invalid", vbCritical, ProgramName$)
            Exit Sub
        End If
        
        ' path would be
        CurrentStore$ = WorkingPath$ & "T4PM_" & ClearSpecialCharacters(GetTempData_WriteBuffer("Project Reference")) & ".xls"
        
        ' ====
        If FileExists(CurrentStore$) = True Then
            Result = MsgBox("A Project Data Store with this reference code already exists!", vbCritical, ProgramName$)
            Exit Sub
        End If
        
        'Adding New Workbook
        Workbooks.Add
        ActiveWorkbook.Sheets(1).Name = "ProjectStore"
        ActiveWorkbook.SaveAs CurrentStore$, xlExcel8
        ActiveWorkbook.Close (True)
        

        Call ExportDataToStore(False)

        Call RefreshRibbon("")

End Sub


