VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectListForm 
   Caption         =   "Project Store Selection"
   ClientHeight    =   4875
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   8310
   OleObjectBlob   =   "ProjectListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub PickProject_Click()

   ' MsgBox WorkingPath$ ' = GetConfigSetting("WorkingPath")

retry:
    ProjectStoreChoice$ = ManualProjectSelection(WorkingPath$)
    
    If VerifyStoreUsers(ProjectStoreChoice$) = False Then
        ProjectStoreChoice$ = ""
        Result = MsgBox("You are not a permitted user for this T4PM Project Store.", vbCritical, ProgramName$)
       ' GoTo retry:
    End If
    
    ProjectListForm.Hide
    
End Sub

Private Sub ProjectStoreList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    On Error Resume Next
    
retry:
    ProjectStoreChoice$ = ProjectStoreList.Column(0)
    If VerifyStoreUsers(ProjectStoreChoice$) = False Then
        ProjectStoreChoice$ = ""
        Result = MsgBox("You are not a permitted user for this T4PM Project Store.", vbCritical, ProgramName$)
        GoTo retry:
    End If
    
    If ProjectStoreChoice$ <> "" Then ProjectListForm.Hide
    
End Sub

Private Sub RefreshList_Click()
    Call DoProjectListRefresh(WorkingPath$)
    
End Sub

Private Sub UserForm_Initialize()
    
        ProjectStoreChoice$ = ""
        

On Error Resume Next

        'only if there is not prog path, refind it
        If ProgramPath$ = "" Then ProgramPath$ = S_UserConfigCode.CheckProgramPath
                
        ' otherwise check if a project list exists
        If FileExists(ProgramPath$ & "ProjectList") = True Then
            ' clear the existig list
            ProjectListForm.ProjectStoreList.Clear

            ' "|||"
            ' open the file
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set a = fs.OpenTextFile(ProgramPath$ & "ProjectList")
                
         ' read line by line
            
            GetAllFile$ = a.readall
            LineCount = (Len(GetAllFile$) - Len(Replace(GetAllFile$, "|||", ""))) / 3 / 4
            
            strdata = Split(GetAllFile$, vbCrLf)
            
            i = 0
            Do Until i >= LineCount
                'strdata = a.readline
                linedata = Split(strdata(i), "|||")
                
                
                ProjectListForm.ProjectStoreList.AddItem linedata(0)
                ProjectListForm.ProjectStoreList.Column(1, i) = linedata(1)
                ProjectListForm.ProjectStoreList.Column(2, i) = linedata(2)
                ProjectListForm.ProjectStoreList.Column(3, i) = linedata(3)
                i = i + 1
            Loop
                
        End If
        
Exit Sub

End Sub
