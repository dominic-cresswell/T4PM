Attribute VB_Name = "PMToolkitCode1"

Public Const ProgramName$ = "GEN² Toolkit for Project Managers (T4PM)"
Public myRibbon
Public RibbonID$

Public ProgramPath$, WorkingPath$
Public RememberProject As Boolean

Public CurrentStore$

Public CurrentDetails$(2)


'Dim FieldListArray(9999, 4)
Public FieldListArray(9999, 4)
Public CollectionListArray(9999)



'============== BUTTON FUNCTIONS

'Callback for customUI.onLoad
Sub Onload(ribbon As IRibbonUI)

    Set myRibbon = ribbon
    
       RibbonID$ = ObjPtr(myRibbon)
        Call SaveRibbonID(RibbonID$)
        

     ' get the field list
      Call ImportFieldData("")
                  
                  

        RememberProject = False
       If PMToolkitCode1.GetConfigSetting("RememberLastProject") <> "" _
           And Replace(LCase(PMToolkitCode1.GetConfigSetting("RememberLastProject")), vbCr, "") = "true" Then
            RememberProject = True
            ' now load the project name from data file
            
            
            ' now load the project itself
       Else
            RememberProject = False
       
       End If
    
        

                  
        
         
End Sub


'======
'Callback for NewProject onAction
Sub NewProject_Click(control As IRibbonControl)

    
    ' check for the four main requirements
    Call PullWriteDataFromWorksheets("")
    
    FailText$ = ""
    
        Test1$ = GetTempData("Site Name")
     If Test1$ = "" Then FailText$ = FailText$ & "Site Name details not known." & vbCrLf
        
        Test2$ = GetTempData("Project Description")
     If Test2$ = "" Then FailText$ = FailText$ & "Project Description details not known." & vbCrLf
    
        Test3$ = GetTempData("Project Manager")
     If Test3$ = "" Then FailText$ = FailText$ & "Project Manager details not known." & vbCrLf
    
        Test4$ = GetTempData("Project Reference")
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
        CurrentStore$ = WorkingPath$ & "T4PM_" & ClearSpecialCharacters(GetTempData("Project Reference")) & ".xls"
        
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



'Callback for PickProject onAction
Sub PickProject_Click(control As IRibbonControl)


 ' check we have the program path to save our array in
    ProgramPath$ = PMToolkitCode1.CheckProgramPath
    
        
 ' get the current working path
    WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))

    PickedFile$ = FindProjectSelection(WorkingPath$)
   
   
   
    If InStr(vbTextCompare, LCase(PickedFile$), "t4pm_") < 1 Then
        Result = MsgBox("Not a T4PM Project Store selected", vbCritical, ProgramName$)
        Exit Sub
    End If
  
  ' =======
    If FileExists(PickedFile$) = False Then
        Result = MsgBox("Invalid Project Store Selection", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    CurrentStore$ = PickedFile$
    
    ' but are we permitted to use this

     Call ClearReadData("")
     Call RestoreStore("")
     Call ImportDataFromStore("")
   '  Call PullWriteDataFromWorksheets("")
     
     '  MsgBox GetTempData2(PermittedUsers_n1)
        
     '   MsgBox "temap 1: " & GetTempData("Project Reference")
     '   MsgBox "temap 2: " & GetTempData2("Project Reference")
        
     '   MsgBox GetAnyDataForHeaders("Project Reference")
        
        Call RefreshRibbon("")
    
    
Exit Sub
abort:
    Result = MsgBox("Invalid Project Store Selection", vbCritical, ProgramName$)

End Sub



'Callback for UploadData onAction
Sub UploadData_Click(control As IRibbonControl)
    'Call Unavailable("")
    
    ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
                Result = MsgBox("Please re-select Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If
        
        
    ' check shift key
    If IsShiftKeyDown = True Then
        Result = MsgBox("Select all worksheets for upload?", vbYesNo, ProgramName$)
        If Result = vbYes Then
            Sheets.Select
        End If
    End If
        
        
    ' re-fill the temp array
    Call PullWriteDataFromWorksheets("")

    ' do the upload
    Call ExportDataToStore(True)

    Call RefreshRibbon("")
    
End Sub

'Callback for DownloadData onAction
Sub DownloadData_Click(control As IRibbonControl)

     ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
                Result = MsgBox("Please re-select Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If


    ' clear the array
    Call ClearReadData("")

    ' get all the info
    Call ImportDataFromStore("")

    ' get the 'type' of data
    
    ' get the grouped data
    
    ' get the special data
    
    ' push it to the field references
        Call PushReadDataToWorksheets("")
        
End Sub



'Callback for WorkingFolderButton onAction
Sub SetFolder_Click(control As IRibbonControl)

    If IsShiftKeyDown = True Then
        Result = MsgBox("Working folder currently set to: " & vbCrLf & vbCrLf & WorkingPath$, vbInformation, ProgramName$)
        Exit Sub
    End If
    
    If WorkingPath$ <> "" And DirExists(WorkingPath$) = True Then
        
        Result = MsgBox("Current Working folder is valid." & vbCrLf & vbCrLf & "Change anyway?", vbInformation + vbYesNo, ProgramName$)
        If Result <> vbYes Then Exit Sub

    End If
    

'
    If CheckUserConfig = False Then ' there is no 'userconfig' file
        Call CreateUserConfig("")
    End If
    
 ' get the current working path
    WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))

    
' select the new working path

    OldWorkingPath$ = WorkingPath$
  On Error GoTo abort:
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
 
        ' Display paths of each file selected
        For lngCount = 1 To 1
            WorkingPath$ = AddSlash(.SelectedItems(lngCount))
        Next lngCount
 
    End With
    
  ' =======
    If DirExists(WorkingPath$) = False Then
        Result = MsgBox("Invalid Folder Selection", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
    GetTxtData$ = Replace(GetTxtData$, OldWorkingPath$, WorkingPath$)
    
    Call MakeTextFile(GetTxtData$, ProgramPath$ & "UserConfigFile")



Exit Sub
abort:
    Result = MsgBox("Invalid Folder Selection", vbCritical, ProgramName$)

End Sub



'======
'Callback for Folder onAction
Sub Folder_Click(control As IRibbonControl)


   ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
               Result = MsgBox("Please re-select Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If
        
        
    ' is shift held... if so, select folder instead of openning
        
        If IsShiftKeyDown = True Then
            
            Result = MsgBox("Force re-selection of Project Folder? ", vbInformation + vbYesNo, ProgramName$)
            If Result = vbYes Then Call SetProjectFolder("")
            
            Exit Sub
        End If
    
    
  'then...
    ' check we have a folder
        
        myFolder$ = PMToolkitCode1.GetAnyDataForHeaders("Folder Path")
        
        If myFolder$ = "" Then
            Result = MsgBox("No Folder Path known." & vbCrLf & "Select now? ", vbCritical + vbYesNo, ProgramName$)
            If Result = vbYes Then Call SetProjectFolder("")
            Exit Sub
        End If
               
        
    ' check the folder is valie
        If DirExists(myFolder$) = False Then
            
            Exit Sub
        End If
        
    
    ' open the folder
        Call OpenFolder(myFolder$)
    
End Sub


'======
'Callback for Email onAction
Sub Email_Click(control As IRibbonControl)


    ' check if we have a store loaded
    
    
    ' check we have a folder & check the folder is valie
    myFolder$ = PMToolkitCode1.GetAnyDataForHeaders("Folder Path")
    If myFolder$ <> "" And DirExists(myFolder$) = True Then Mail_savepath$ = myFolder$
       

    ' create an email
    Mail_subject$ = ""
    Mail_subject$ = Mail_subject$ & PMToolkitCode1.GetAnyDataForHeaders("Site Name") & " - "
    Mail_subject$ = Mail_subject$ & PMToolkitCode1.GetAnyDataForHeaders("Project Description")
    Mail_subject$ = Mail_subject$ & " (" & PMToolkitCode1.GetAnyDataForHeaders("Project Reference") & ")"

   ' Mail_savepath$ = "C:"
    Call NewMail

    
End Sub


Sub RecallProject_Status(control As IRibbonControl, ByRef returnedVal)
'
' Code for getPressed callback. Ribbon control checkBox
'

    'If control.ID = "checkboxShowMessage" Then
        returnedVal = RememberProject
    'End If
    
    
    
End Sub

Sub RecallProject_Click(control As IRibbonControl, pressed As Boolean)
'
' Code for onAction callback. Ribbon control checkBox
'
    If control.ID = "checkboxShowMessage" Then
       RememberProject = pressed
    End If
    
    
    Call SetConfigSetting("RememberLastProject", CStr(pressed))
    
    
    
End Sub


Sub SetConfigSetting(inOption$, inParam$)

    If ProgramPath$ = "" Then ProgramPath$ = PMToolkitCode1.CheckProgramPath
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
 
    ' check for old
    
    If GetConfigSetting(inOption$) <> "" Then
        OldSetting$ = inOption$ + "=" + GetConfigSetting(inOption$)
        
        GetTxtData$ = Replace(GetTxtData$, OldSetting$, inOption$ + "=" + inParam$)
        

    Else
        GetTxtData$ = GetTxtData$ & inOption$ + "=" + inParam$ & vbCrLf
    End If
    
    
    Call MakeTextFile(GetTxtData$, ProgramPath$ & "UserConfigFile")
    

End Sub


Function GetConfigSetting(inOption$)
    
    ' ....
If ProgramPath$ = "" Then ProgramPath$ = PMToolkitCode1.CheckProgramPath
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
 
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
  
       GetPoint = InStr(vbTextCompare, LCase(GetTxtData$), LCase(inOption$) & "=")
    If GetPoint = 0 Then Exit Function
    
    GetConfigSetting = Right(GetTxtData$, Len(GetTxtData$) - GetPoint - Len(inOption$))
    GetPoint = InStr(vbTextCompare, GetConfigSetting, vbCrLf)
    
    
'====
    If GetPoint > 0 Then
        GetConfigSetting = Left(GetConfigSetting, GetPoint - 1)
    End If
        
    
End Function


Function AddSlash(inString$) As String

    AddSlash = inString$
         If Right(AddSlash, 1) <> ":" And AddSlash <> "" And Right(AddSlash, 1) <> "\" Then
         AddSlash = AddSlash & "\"
        End If

   
End Function



Sub CreateUserConfig(dummy$)

        DefaultConfig$ = ""
        DefaultConfig$ = DefaultConfig$ & "WorkingPath=" & Environ("userprofile") & "\" & vbCrLf
        DefaultConfig$ = DefaultConfig$ & "RememberLastProject=False" & vbCrLf
        DefaultConfig$ = DefaultConfig$ & "" & vbCrLf
     
   Call MakeTextFile(DefaultConfig$, ProgramPath$ & "UserConfigFile")


End Sub

'Callback for GetFieldListButton onAction
Sub GetList_Click(control As IRibbonControl)


 ' check we have the program path to save our array in
    ProgramPath$ = PMToolkitCode1.CheckProgramPath

 ' get the current working path
        WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))
    
 ' lets pick the File (FieldReferences)
    
   If FileExists(WorkingPath$ & "FieldReferences.xlsx") Then
        FieldRefFile$ = WorkingPath$ & "FieldReferences.xlsx"
    
   Else
    
  ' select the new working path
         FieldRefFile$ = ""
         On Error GoTo abort:
           With Application.FileDialog(msoFileDialogFilePicker)
               .AllowMultiSelect = False
               .Filters.Clear
               .Filters.Add "Excel Workbook", "*.xls*", 1
               .Show
               
               ' Display paths of each file selected
               For lngCount = 1 To 1
                   FieldRefFile$ = .SelectedItems(lngCount)
               Next lngCount
        
           End With
    End If
    

    
    
skipselector:

    ' check if it's right
    If FieldRefFile$ = "" Or InStr(vbTextCompare, LCase(FieldRefFile$), ".xls") < 0 Then GoTo abort:
    
    Dim FieldExc As Workbook
    On Error GoTo abort:
    Set FieldExc = Application.Workbooks.Open(FieldRefFile$, False, True)
    

    ' clear the fieldlist table
    
      For XXX = 0 To 4
         For YYY = 0 To 9999
            FieldListArray(YYY, XXX) = ""
         Next
        Next
        
    
'    open the fieldlist table
        Dim mysheet As Worksheet
           
        On Error GoTo closeabort:
        Set mysheet = FieldExc.Worksheets("FieldList")
         If mysheet Is Nothing Then GoTo closeabort:
    
       ' mop up all the FieldListArray(9999, 4)
       counter = 0
       ' For XXX = 1 To 5
         For YYY = 1 To 9999
        
           On Error Resume Next
            DataRef$ = ""
            DataDescription$ = ""
            DataType$ = ""
            DataCollection$ = ""
            DataMultiplier = False
            
                    DataRef$ = mysheet.Columns(1).Rows(YYY)
            DataDescription$ = mysheet.Columns(2).Rows(YYY)
                   DataType$ = LCase(mysheet.Columns(3).Rows(YYY))
             DataCollection$ = mysheet.Columns(4).Rows(YYY)
              DataMultiplier = CBool(mysheet.Columns(5).Rows(YYY))
    
        ' check that we are declaring a datatype
           If DataType$ = "text" Or _
              DataType$ = "numerical" Or _
              DataType$ = "memo" Or _
              DataType$ = "boolean" Or _
              DataType$ = "date" Then
     
                FieldListArray(counter, 0) = DataRef$
                FieldListArray(counter, 1) = DataDescription$
                FieldListArray(counter, 2) = DataType$
                FieldListArray(counter, 3) = DataCollection$
                FieldListArray(counter, 4) = DataMultiplier
                
                    counter = counter + 1

           End If
         Next
       ' Next XXX
    
    FieldExc.Close
    
    On Error GoTo abort:
    

    ' save out the data
         
    MajorFieldData$ = ""
   On Error GoTo 0
   
    For YYY = 0 To 9999
    For XXX = 0 To 4
      MajorFieldData$ = MajorFieldData$ & FieldListArray(YYY, XXX) & Chr(255)
    Next
    Next

    Call MakeTextFile(MajorFieldData$, ProgramPath$ & "FieldData")
    
    Result = MsgBox("Field List Updated.", vbInformation, ProgramName$)


    
    Exit Sub
    
closeabort:
    FieldExc.Close
    
abort:
    Result = MsgBox("Invalid Field List Selection", vbCritical, ProgramName$)
    
End Sub


Sub IssueSheet_Click(control As IRibbonControl)

    Call MakePDFSheet("")

End Sub


Sub Unavailable(dummy$)

    Result = MsgBox("Function not yet available.", vbCritical, ProgramName$)
    
End Sub


Private Function GetAnyDataForHeaders(inString$) As String

If inString$ = "" Then Exit Function

        ' check the 'write' (updated) data
    '   GetAnyDataForHeaders = CStr(GetTempData(inString$))
       
       ' check the 'read' (stored) data
       If GetAnyDataForHeaders = "" Then
       If ProjectReadDataArray(0, 0) = "" Then Call ImportDataFromStore("")
       GetAnyDataForHeaders = GetTempData2(inString$)
       End If
  
End Function
       
       
Sub CallbackGetSiteLabel(control As IRibbonControl, ByRef label)

     label = " Site: ___________________________________________________"
    
       FindSite$ = CStr(GetAnyDataForHeaders("Site Name"))
    If FindSite$ = "" Then Exit Sub
        
        If Len(FindSite$) > 48 Then FindSite$ = Left(FindSite$, 48) & "..."
           label = " Site:  " & FindSite$
 
End Sub

Sub CallbackGetTitleLabel(control As IRibbonControl, ByRef label)

     label = "Title: ___________________________________________________"

       FindProjectTitle$ = CStr(GetAnyDataForHeaders("Project Description"))
    If FindProjectTitle$ = "" Then Exit Sub
        
        If Len(FindProjectTitle$) > 48 Then FindProjectTitle$ = Left(FindProjectTitle$, 48) & "..."
        label = "Title:  " & FindProjectTitle$
 
End Sub

Sub CallbackGetReferenceLabel(control As IRibbonControl, ByRef label)

     label = " Ref.: ___________________________________________________"
    
    FindProjectRef$ = CStr(GetAnyDataForHeaders("Project Reference"))
       
    If FindProjectRef$ = "" Then Exit Sub
        
        If Len(FindProjectRef$) > 48 Then FindProjectRef$ = Left(FindProjectRef$, 48) & "..."
        label = " Ref.:  " & FindProjectRef$

End Sub




'----------------

Private Function CheckUserConfig() As Boolean
        
    CheckUserConfig = False
       ProgramPath$ = PMToolkitCode1.CheckProgramPath

' check the ...
    If FileExists(ProgramPath$ & "UserConfigFile") = False Then
            CheckUserConfig = False
    Else
            CheckUserConfig = True
    End If
     
     
End Function
    
    
Public Function CheckProgramPath() As String

' = = =
       ConfigFile$ = Environ("appdata")
        
    If Right(ConfigFile$, 1) <> "\" Then ConfigFile$ = ConfigFile$ & "\"

' .. set up the sub-folder
    If DirExists(ConfigFile$ & "T4PM") = False Then
        MkDir ConfigFile$ & "T4PM"
    End If
    
    CheckProgramPath = ConfigFile$ & "T4PM"
    If Right(CheckProgramPath, 1) <> "\" Then CheckProgramPath = CheckProgramPath & "\"
    
End Function



Sub ImportFieldData(dummy$)

   ' but we must clear the list first!!
    For rrr = 0 To 9999
        For eee = 0 To 4
            FieldListArray(rrr, eee) = ""
        Next
    Next

    MajorFieldData$ = ""
  
    ' routine check
    ProgramPath$ = PMToolkitCode1.CheckProgramPath
        
    ' field data exists?
    If FileExists(ProgramPath$ & "FieldData") = False Then
        Result = MsgBox("No Field Data available locally." & vbCrLf & vbclr & "Please re-import.", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    'load data
    MajorFieldData$ = ReadTextFile(ProgramPath$ & "FieldData")
    

    
    ' process data
     
     For YYY = 0 To 9999
     For XXX = 0 To 4
     FieldListArray(YYY, XXX) = ""
     Next
     Next
    
    XXX = 0
    YYY = 0
    For Z = 1 To Len(MajorFieldData$)
        CheckMe$ = Mid(MajorFieldData$, Z, 1)
        If Asc(CheckMe$) <> 255 Then
            GetData$ = GetData$ & CheckMe$
        Else
            
            'Debug.Print GetData$
            
            
            FieldListArray(YYY, XXX) = GetData$
            XXX = XXX + 1
            If XXX = 5 Then
                XXX = 0
                YYY = YYY + 1
            End If
            If YYY >= 9999 Then
                Exit For
            End If
            
            GetData$ = ""
        End If
    Next
    
End Sub


