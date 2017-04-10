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
    ProjectStoreChoice$ = ManualProjectSelection(WorkingPath$)
    ProjectListForm.Hide
    
End Sub

Private Sub ProjectStoreList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    On Error Resume Next
    ProjectStoreChoice$ = ProjectStoreList.Column(0)
    
    If ProjectStoreChoice$ <> "" Then ProjectListForm.Hide
    
End Sub

Private Sub RefreshList_Click()
    Call DoProjectListRefresh(WorkingPath$)
    
End Sub

Private Sub UserForm_Initialize()
    
        ProjectStoreChoice$ = ""

      '  Stop
End Sub
