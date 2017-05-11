Attribute VB_Name = "PMToolkitCode5"

Public ProjectStoreChoice$






Function FindProjectSelection(inPath$)

    ProjectListForm.Show
    FindProjectSelection = ProjectStoreChoice$
    Debug.Print " selected: " & FindProjectSelection
    
    
 ' routine check
    ProgramPath$ = AddSlash(PMToolkitCode1.CheckProgramPath)
    
   ' +++++++++++
    If RememberProject = True Then
        ' save project name
        Call MakeTextFile(ProjectStoreChoice$, ProgramPath$ & "LastProject")

  ' +++++++++++
    Else
        If FileExists(ProgramPath$ & "LastProject") Then Kill ProgramPath$ & "LastProject"
        
    End If
    
    
End Function



Function ManualProjectSelection(inPath$) As String

    inPath$ = AddSlash(inPath$)

' select the new working path
  On Error GoTo abort:
  
    With Application.FileDialog(msoFileDialogFilePicker)
         .InitialFileName = inPath$
         .Filters.Clear
         .Filters.Add "T4PM Excel Files", "*.xls*", 1
    
         .AllowMultiSelect = False
         .Show
 
        ' Display paths of each file selected
        For lngCount = 1 To 1
            ManualProjectSelection = (.SelectedItems(lngCount))
        Next lngCount
        
        ' ++++
        
        
 
    End With
    

 
abort:
End Function


Sub DoProjectListRefresh(inPath$)

    
    inPath$ = AddSlash(inPath$)
        
    ' scroll through all items

        ProjectListData$ = ""
       
        i = 0
        flag = True
        File = Dir(inPath$, vbNormal)
        ProjectListForm.ProjectStoreList.Clear
        
        While File <> ""
                If (Right(File, 4) = ".xls" Or Right(File, 5) = ".xlsm" Or Right(File, 5) = ".xlsx") And Left(File, 5) = "T4PM_" Then
                
                    Debug.Print File

                    Dim prjd As ProjectData
                    prjd = GetTopData(inPath$ & File)
                    
                    Debug.Print prjd.SiteName
                    Debug.Print prjd.ProjectDescription
                    Debug.Print prjd.ProjectReference
                    Debug.Print prjd.AllUsers
                    
                    If prjd.SiteName <> "" And prjd.ProjectReference <> "" And InStr(vbTextCompare, LCase(prjd.AllUsers), LCase(Environ("UserName"))) > 0 Then
                        
                        ProjectListForm.ProjectStoreList.AddItem (inPath$ & File)
                        ProjectListForm.ProjectStoreList.Column(1, i) = CapText(prjd.SiteName, 38)
                        ProjectListForm.ProjectStoreList.Column(2, i) = CapText(prjd.ProjectDescription, 38)
                        ProjectListForm.ProjectStoreList.Column(3, i) = prjd.ProjectReference
                        
                        ProjectListData$ = ProjectListData$ & "" & (inPath$ & File)
                        ProjectListData$ = ProjectListData$ & "|||" & CapText(prjd.SiteName, 38)
                        ProjectListData$ = ProjectListData$ & "|||" & CapText(prjd.ProjectDescription, 38)
                        ProjectListData$ = ProjectListData$ & "|||" & prjd.ProjectReference
                        ProjectListData$ = ProjectListData$ & "|||" & vbCrLf
                        
                        i = i + 1
                    End If
                
                    
                End If
                File = Dir
        'Stop
        Wend
                
        'Stop
        
    ' check we have the program path to save our array in
         If ProgramPath$ = "" Then ProgramPath$ = PMToolkitCode1.CheckProgramPath
        ' save list to user folder
        If ProgramPath$ <> "" And DirExists(ProgramPath$) Then
         Call MakeTextFile(ProjectListData$, ProgramPath$ & "ProjectList")
        End If
        
End Sub

Function CapText(inText$, amount As Long) As String

    CapText = inText$
    If Len(CapText) > amount Then CapText = Left(CapText, amount - 3) & "..."

End Function



Function GetTopData(inFile$) As ProjectData
    Dim prjd As ProjectData
    
    prjd.SiteName = ""
    prjd.ProjectDescription = ""
    prjd.ProjectReference = ""
    prjd.AllUsers = ""


' invoke a new Excel
    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")
    exlApp.visible = False
    
    If inFile$ = "" Then Exit Function
    
       
    Dim exlDoc As Workbook
    On Error GoTo fail2:
    Set exlDoc = exlApp.Workbooks.Open(inFile$)

    
    Dim exlSheet As Worksheet
    On Error Resume Next
    Set exlSheet = exlDoc.Worksheets.Item("ProjectStore")
    
    If exlSheet Is Nothing Then GoTo fail:


   ' we made it!

      ' find existing data
        For QQQ = 1 To 9999

                FieldName$ = exlSheet.Columns(1).Rows(QQQ)
                FieldData$ = exlSheet.Columns(2).Rows(QQQ)
            '    FieldStamp$ = exlSheet.Columns(3).Rows(qqq)
                
                If FieldName$ = "" Then Exit For
                
             '   ProjectReadDataArray(zzz, 0) = FieldName$
             '   ProjectReadDataArray(zzz, 1) = FieldData$
               ' ProjectReadDataArray(zzz, 2) = FieldStamp$
                
                
                If FieldName$ = "SiteName_n0" Then prjd.SiteName = FieldData$
                If FieldName$ = "ProjectDescription_n0" Then prjd.ProjectDescription = FieldData$
                If FieldName$ = "ProjectReference_n0" Then prjd.ProjectReference = FieldData$
                
                If Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers = "" Then
                    prjd.AllUsers = FieldData$
                ElseIf Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers <> "" Then
                    prjd.AllUsers = prjd.AllUsers & ", " & FieldData$
                End If

                zzz = zzz + 1
        Next

    exlDoc.Close (False)
    exlApp.Quit
    
    
    GetTopData = prjd
    
  '  If showmsg = True Then Result = MsgBox("Data Downloaded", vbInformation, ProgramName$)
    Exit Function
    
    

    
    
    

    
    
    

fail:
    exlDoc.Close
fail2:
    exlApp.Quit
    'result = MsgBox("No worksheet 'Project Store' within working store.", vbCritical, ProgramName$)
        
    GetTopData = prjd
    
End Function

