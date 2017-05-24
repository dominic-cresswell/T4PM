Attribute VB_Name = "E_RibbonEditorCode"
Option Private Module


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


