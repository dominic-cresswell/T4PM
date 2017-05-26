Attribute VB_Name = "S_RibbonButtonCode"
Option Private Module


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
    
' ~~~~~~~~~~~~~~~~~~
        If S_UserConfigCode.GetConfigSetting("RememberLastProject") <> "" _
           And Replace(LCase(S_UserConfigCode.GetConfigSetting("RememberLastProject")), vbCr, "") = "true" Then
            RememberProject = True
           
          ' one stop shop for finding remembered projects
            Call RestoreStore("")
            
       Else
            RememberProject = False
    
       End If
    
    
    
'   If Application.Name = "Microsoft Excel" Then Call E_RibbonButtonCode.ExcelRibbonLoad
         
End Sub


'======



'================ +++ SETTINGS GROUP


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


'Callback for GetFieldListButton onAction
Sub GetList_Click(control As IRibbonControl)


 ' check we have the program path to save our array in
    ProgramPath$ = S_UserConfigCode.CheckProgramPath

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
    Set FieldExc = Excel.Workbooks.Open(FieldRefFile$, False, True)
    

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




Sub RecallProject_Status(control As IRibbonControl, ByRef returnedVal)

    'If control.ID = "checkboxShowMessage" Then
        returnedVal = RememberProject
    'End If
    
    
End Sub

Sub RecallProject_Click(control As IRibbonControl, pressed As Boolean)

    If control.ID = "checkboxShowMessage" Then
       RememberProject = pressed
    End If
    
    Call SetConfigSetting("RememberLastProject", CStr(pressed))
    
    
End Sub



' =+++++++++++++++++++++++++= DATA TOOLS GROUP


'Callback for PickProject onAction
Sub PickProject_Click(control As IRibbonControl)

 ' check we have the program path to save our array in
    ProgramPath$ = S_UserConfigCode.CheckProgramPath
    
 ' get the current working path
    WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))
    PickedFile$ = FindProjectSelection(WorkingPath$)
   
  '
    LoadProjectStore (PickedFile$)

End Sub



'Callback for UploadData onAction
Sub UploadData_Click(control As IRibbonControl)
    If Application.Name <> "Microsoft Excel" Then
        Result = MsgBox("Function currently only supported on Microsoft Excel", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
               Result = MsgBox("Please re-select T4PM Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If
        
        
    ' check reference numbers!
      Dim prjd As ProjectData
          prjd = GetTopData(CurrentStore$)
          ActiveRef$ = prjd.ProjectReference
          LiveRef$ = GetLiveReferenceCode
          
          
      If LiveRef$ = "" Then
            Result = MsgBox("The Active Workbook does not have a completed Reference Number field.", vbCritical, ProgramName$)
            Exit Sub
      ElseIf LiveRef$ <> ActiveRef$ Then
            Result = MsgBox("Active Workbook Reference Number (" & LiveRef$ & ") does not match selected T4PM Project Store. (" & ActiveRef$ & ")", vbCritical, ProgramName$)
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
                Result = MsgBox("Please re-select T4PM Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If


    ' clear the array
    Call ClearReadData("")

    ' get all the info
    Call ImportDataFromStore("")

    ' get the 'type' of data (include collection/individual etc)
    
    ' get the grouped data
    
    ' get the special data
    
    ' push it to the field references
       If Application.Name = "Microsoft Excel" Then Call E_DataUploadCode.PushReadDataToWorksheets("")
       If Application.Name = "Microsoft Word" Then Call W_DownloadCode.PushReadDataToDocument("")
       
End Sub


Sub IssueSheet_Click(control As IRibbonControl)

    Call MakePDFSheet("")

End Sub


'======
'Callback for Email onAction
Sub Email_Click(control As IRibbonControl)


    ' check if we have a store loaded
    
    
    ' check we have a folder & check the folder is valie
    myFolder$ = S_DownloadCode.GetAnyDataForHeaders("Folder Path")
    If myFolder$ <> "" And DirExists(myFolder$) = True Then Mail_savepath$ = myFolder$
       

    ' create an email
    Mail_subject$ = ""
    Mail_subject$ = Mail_subject$ & S_DownloadCode.GetAnyDataForHeaders("Site Name") & " - "
    Mail_subject$ = Mail_subject$ & S_DownloadCode.GetAnyDataForHeaders("Project Description")
    Mail_subject$ = Mail_subject$ & " (" & S_DownloadCode.GetAnyDataForHeaders("Project Reference") & ")"

   ' Mail_savepath$ = "C:"
    Call NewMail

    
End Sub




'================ +++ CURRENT PROJECT GROUP


'======
'Callback for Folder onAction
Sub Folder_Click(control As IRibbonControl)


   ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
               Result = MsgBox("Please re-select T4PM Project Store", vbCritical, ProgramName$)
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
        
        myFolder$ = S_DownloadCode.GetAnyDataForHeaders("Folder Path")
        
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

