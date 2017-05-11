Attribute VB_Name = "PMToolkitCode4"

Sub RestoreStore(dummy$)

    ' grab the working path
            
        If WorkingPath$ = "" Then
            WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))
        End If
        
        If WorkingPath$ <> "" And DirExists(WorkingPath$) = True Then
            TempStore$ = WorkingPath$
        
        End If
        
    ' grab the reference number
     If GetTempData("Project Reference") = "" Then
         Call PullWriteDataFromWorksheets("")
     End If
     
     
     TempRef$ = GetTempData("Project Reference")
            
     If TempRef$ <> "" And TempStore$ <> "" Then
         TempStore$ = TempStore$ & TempRef$
     End If
     
     
    ' check file exists
    If FileExists("T4PM" & TempStore$ & ".xls") = False Then
            TempStore$ = ""
    End If

    If TempStore$ <> "" Then CurrentStore$ = "T4PM" & TempStore$ & ".xls"

End Sub


Sub ExportDataToStore(showmsg As Boolean)

    ' ======
    If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
        Result = MsgBox("No Project Store selected.", vbCritical, ProgramName$)
        Exit Sub
    End If

    ' ======
    If ProjectWriteDataArray(zzz, 0) = "" Then
    Call PullWriteDataFromWorksheets("")
    End If
    
    
   ' invoke a new Excel
    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")
    exlApp.visible = False
       
    Dim exlDoc As Workbook
    Set exlDoc = exlApp.Workbooks.Open(CurrentStore$)

    
    
    Dim exlSheet As Worksheet
    On Error Resume Next
    Set exlSheet = exlDoc.Worksheets.Item("ProjectStore")
    
    If exlSheet Is Nothing Then GoTo fail:
    

    
    ' we made it!
    For zzz = 0 To 9999
    
        FieldName$ = ProjectWriteDataArray(zzz, 0)
        FieldData$ = ProjectWriteDataArray(zzz, 1)
        
        If FieldName$ = "" Then Exit For
        
        ' find a spot
        For QQQ = 1 To 9999
            If (exlSheet.Columns(1).Rows(QQQ) = "" Or LCase(exlSheet.Columns(1).Rows(QQQ)) = LCase(FieldName$)) _
                And FieldName$ <> "" Then
                
                ' check the reference code hasnt changed!
                If LCase(exlSheet.Columns(1).Rows(QQQ)) = "projectreference_n0" And (exlSheet.Columns(2).Rows(QQQ) <> FieldData$) Then
                    Result = MsgBox("Reference Number change has not been stored.", vbCritical, ProgramName$)
                Else
                
                    exlSheet.Columns(1).Rows(QQQ) = FieldName$
                    exlSheet.Columns(2).Rows(QQQ) = FieldData$
                    exlSheet.Columns(3).Rows(QQQ) = Format(Date + Time, "dd-mmm-yyyy hh:mm")
                End If
                
                Exit For
            
            End If
        Next
        
        '===
    
    Next
    
    exlApp.DisplayAlerts = False
    exlDoc.Save
    exlApp.DisplayAlerts = True
    exlDoc.Close (True)
    exlApp.Quit
    
    Call ImportDataFromStore("")
    
    If showmsg = True Then Result = MsgBox("Data Uploaded", vbInformation, ProgramName$)
    Exit Sub
    
    
fail:
    exlDoc.Close
    exlApp.Quit
    Result = MsgBox("No worksheet 'Project Store' within working store.", vbCritical, ProgramName$)
    
End Sub


Sub ImportDataFromStore(dummy$)

' invoke a new Excel
    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")
    exlApp.visible = False
    
    If CurrentStore$ = "" Then Exit Sub
    
       
    Dim exlDoc As Workbook
    On Error GoTo fail2:
    Set exlDoc = exlApp.Workbooks.Open(CurrentStore$)

    
    Dim exlSheet As Worksheet
    On Error Resume Next
    Set exlSheet = exlDoc.Worksheets.Item("ProjectStore")
    
    If exlSheet Is Nothing Then GoTo fail:


   ' we made it!

      ' find existing data
        For QQQ = 1 To 9999

                FieldName$ = exlSheet.Columns(1).Rows(QQQ)
                FieldData$ = exlSheet.Columns(2).Rows(QQQ)
                FieldStamp$ = exlSheet.Columns(3).Rows(QQQ)
                
                If FieldName$ = "" Then Exit For
                
                ProjectReadDataArray(zzz, 0) = FieldName$
                ProjectReadDataArray(zzz, 1) = FieldData$
                ProjectReadDataArray(zzz, 2) = FieldStamp$
                
                zzz = zzz + 1
        Next



    exlDoc.Close (False)
    exlApp.Quit
    
  '  If showmsg = True Then Result = MsgBox("Data Downloaded", vbInformation, ProgramName$)
    Exit Sub
    
    
fail:
    exlDoc.Close
fail2:
    exlApp.Quit
    Result = MsgBox("No worksheet 'Project Store' within working store.", vbCritical, ProgramName$)
    

End Sub

Sub ClearReadData(dummy$)

    For zzz = 0 To 9999
     For QQQ = 0 To 4
        ProjectReadDataArray(zzz, QQQ) = ""
     Next
    Next
    
End Sub


Sub ClearWriteData(dummy$)

    For zzz = 0 To 9999
     For QQQ = 0 To 4
        ProjectWriteDataArray(zzz, QQQ) = ""
     Next
    Next
    
End Sub


Sub PushReadDataToWorksheets(dummy$)
    

    For zzz = 0 To 9999
            FieldName1$ = ProjectReadDataArray(zzz, 0)
            
            If FieldName1$ = "" Then Exit For
            
            FieldName1$ = Replace(FieldName1$, "_n0", "_null")
            FieldName2$ = FieldName1$
            
            FieldName1$ = "T4PM_S_W_" & FieldName1$
            FieldName2$ = "T4PM_S_R_" & FieldName2$
            
            FieldData$ = ProjectReadDataArray(zzz, 1)
            FieldStamp$ = ProjectReadDataArray(zzz, 2)
            
            For Each sh In Sheets
            
                On Error Resume Next

                sh.Range(FieldName1$) = FieldData$
                sh.Range(FieldName2$) = FieldData$
            
            Next
            
    Next
    
End Sub


Sub ClearDataInWorkbook(dummy$)

  ' check shift key
    If IsShiftKeyDown = True Then
        Result = MsgBox("Select all worksheets for upload?", vbYesNo, ProgramName$)
        If Result = vbYes Then
            Sheets.Select
        End If
    End If
    
    On Error Resume Next
    
    
    Result = MsgBox("This function will remove all current data within the Workbook." & vbCrLf & vbCrLf & "This cannot be undone." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbInformation, ProgramName$)
    If Result <> vbYes Then Exit Sub
    
    On Error Resume Next
    
    
   '         For Each sh In Sheets
            For Each sh In Excel.ActiveWindow.SelectedSheets
                For Each subName In sh.Names
                  
                'Debug.Print LCase(Left(Replace(subName.Name, sh.Name & "!", ""), 5))
                
                realRangeName = subName.Name
                If Left(subName.Name, 1) = "'" Then realRangeName = Replace(realRangeName, "'" & sh.Name & "'!", "")
                If Left(subName.Name, 1) <> "'" Then realRangeName = Replace(realRangeName, sh.Name & "!", "")
                
                  If Left(LCase(realRangeName), 5) = "t4pm_" Then
                  
                    sh.Range(subName.Name).Cells(1) = ""
                  End If
                  
                Next
            Next
    
End Sub


Sub DeleteDynamicRange(dummy$)

    ' check we have a selection
    If Selection.Count < 1 Then Exit Sub

    ' confirm with user
       Result = MsgBox("This function will remove any Dynamic Fields from the currently selected cell(s)." & vbCrLf & vbCrLf & "This cannot be undone." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbInformation, ProgramName$)
    If Result <> vbYes Then Exit Sub

    On Error Resume Next
            
    ' check all selected cells (loop)
    For Each bitcell In Selection.Cells

            ' in every sheet
            For Each sh In Sheets
            
            ' look through all the ranges
            For Each rangeNames In sh.Names
            
            ' get the cell (selection) and the named range
            Set testCell = sh.Range(bitcell.Address)
            Set testRange = sh.Range(rangeNames.Name)
            
            ' test them

            Set isect = Application.Intersect(bitcell, testRange)
        
        
            If isect Is Nothing Then
            Else
                ' we have a selection
                isect.Select
            '    MsgBox "Range " & bitcell.Address & " DOES  intersect with " & rangeNames.Name
                     
               ' we have an interect. Check it's a T4PM one.
                If Left(Replace(rangeNames.Name, sh.Name & "!", ""), 5) = "T4PM_" Then
                rangeNames.Delete
                End If
                 
            End If
            Next
            Next
    Next

End Sub
