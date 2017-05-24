Attribute VB_Name = "S_ProjectListFormCode"
Option Private Module

Public ProjectStoreChoice$



Function FindProjectSelection(inPath$)

    ProjectListForm.Show
    FindProjectSelection = ProjectStoreChoice$
    Debug.Print " selected: " & FindProjectSelection
    
    
 ' routine check
    ProgramPath$ = AddSlash(S_UserConfigCode.CheckProgramPath)
    
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
         If ProgramPath$ = "" Then ProgramPath$ = S_UserConfigCode.CheckProgramPath
         
      ' save list to user folder
        If ProgramPath$ <> "" And DirExists(ProgramPath$) Then
         Call MakeTextFile(ProjectListData$, ProgramPath$ & "ProjectList")
        End If
        
End Sub



Function VerifyStoreUsers(inFile$) As Boolean


    If FileExists(inFile$) = False Or inFile$ = "" Then Exit Function

    If IsShiftKeyDown = True Then

            Password$ = InputBoxDK("Enter override password", ProgramName$)
            If Password$ <> "onetwothree" Then
                VerifyStoreUsers = False
                Result = MsgBox("Password incorrect", vbCritical, ProgramName$)
            Else
                VerifyStoreUsers = True
                Exit Function
            End If
            
    End If
    
    
       Dim prjd As ProjectData
        prjd = GetTopData(inFile$)
        
        Debug.Print "file: " & inFile$
        Debug.Print "users: " & prjd.AllUsers
        
        If InStr(vbTextCompare, LCase(prjd.AllUsers), LCase(Environ("UserName"))) > 0 Then
            VerifyStoreUsers = True
        Else
            VerifyStoreUsers = False
        End If
        
End Function



