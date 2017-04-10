Attribute VB_Name = "PMToolkitCode5"

Public ProjectStoreChoice$






Function FindProjectSelection(inPath$)

    ProjectListForm.Show
    FindProjectSelection = ProjectStoreChoice$
    Debug.Print " selected: " & FindProjectSelection
    
    
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

       
        I = 0
        flag = True
        file = Dir(inPath$, vbNormal)
        ProjectListForm.ProjectStoreList.Clear
        
        While file <> ""
                If (Right(file, 4) = ".xls" Or Right(file, 5) = ".xlsm" Or Right(file, 5) = ".xlsx") And Left(file, 5) = "T4PM_" Then
                
                    Debug.Print file

                    Dim prjd As ProjectData
                    prjd = GetTopData(inPath$ & file)
                    
                    Debug.Print prjd.SiteName
                    Debug.Print prjd.ProjectDescription
                    Debug.Print prjd.ProjectReference
                    Debug.Print prjd.AllUsers
                    
                    If prjd.SiteName <> "" And prjd.ProjectReference <> "" And InStr(vbTextCompare, LCase(prjd.AllUsers), LCase(Environ("UserName"))) > 0 Then
                        
                        ProjectListForm.ProjectStoreList.AddItem (inPath$ & file)
                        ProjectListForm.ProjectStoreList.Column(1, I) = CapText(prjd.SiteName, 38)
                        ProjectListForm.ProjectStoreList.Column(2, I) = CapText(prjd.ProjectDescription, 38)
                        ProjectListForm.ProjectStoreList.Column(3, I) = prjd.ProjectReference
                        I = I + 1
                    End If
                    
                    
                End If
                file = Dir
        'Stop
        Wend
                
        'Stop


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
        For qqq = 1 To 9999

                FieldName$ = exlSheet.Columns(1).Rows(qqq)
                FieldData$ = exlSheet.Columns(2).Rows(qqq)
            '    FieldStamp$ = exlSheet.Columns(3).Rows(qqq)
                
                If FieldName$ = "" Then Exit For
                
             '   ProjectReadDataArray(zzz, 0) = FieldName$
             '   ProjectReadDataArray(zzz, 1) = FieldData$
               ' ProjectReadDataArray(zzz, 2) = FieldStamp$
                
                
                If FieldName$ = "SiteName_n0" Then prjd.SiteName = FieldData$
                If FieldName$ = "ProjectDescription_n0" Then prjd.ProjectDescription = FieldData$
                If FieldName$ = "ProjectReference_n0" Then prjd.ProjectReference = FieldData$
                
                If Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers = "" Then prjd.AllUsers = FieldData$
                If Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers <> "" Then prjd.AllUsers = prjd.AllUsers & ", " & FieldData$


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

