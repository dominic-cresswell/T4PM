Attribute VB_Name = "W_RibbonEditorButtonCode"
Option Private Module



Public Const ApplicationVersion = 1

' ======================= BUTTONS

'Callback for SetFieldButton onAction
Sub SetTemplateField_Click(control As IRibbonControl)

    FieldRefOutput$ = ""

' check selection (you'll need to have a word document open)
'    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveDocument Is Nothing Then Exit Sub
    
    ' check selection
        If Selection.Active = False Then
        Result = MsgBox("No area selected.", vbCritical, ProgramName$)
        Exit Sub
        End If

    ' get field data
    Call SetupField("")


'    ' nothing selected
    If FieldRefOutput$ <> "" Then
    Else
        Result = MsgBox("No Dynamic Field selected", vbCritical, ProgramName$)
        Exit Sub
    End If


'   Result = MsgBox("This Dynamic Field already in " & vbCrLf & "use on this worksheet." & vbCrLf & vbCrLf & "Delete Previous?", vbCritical + vbYesNo, ProgramName$)

     If Selection.ShapeRange.Count > 0 Then

        For Each shp In Selection.ShapeRange
            shp.TextFrame.TextRange.Text = "<<" & Replace(FieldRefOutput$, "_W_", "_R_") & ">>"
        Next
     
     Else

       Selection.TypeText Text:="<<" & Replace(FieldRefOutput$, "_W_", "_R_") & ">>"
                
     End If


    
End Sub




Sub SetupField(dummy$)

    ' if the array is empty, reload it
    If CStr(FieldListArray(0, 1)) = "" Then
        Call ImportFieldData("")
    End If

    ' if the array is *still* empty, tell the user to import again
    If CStr(FieldListArray(0, 1)) = "" Then
       Exit Sub
    End If
    
  ' we made it here, we must be OK.
  

  '
    Call RefreshFieldList("")
            
    FieldRefForm.Show
        
End Sub

'Sub RemoveFieldButton_Click(control As IRibbonControl)
'
'    Call DeleteDynamicRange("")
'
'End Sub

'Sub ClearFieldButton_Click(control As IRibbonControl)'
'
'    Call ClearDataInWorkbook("")'
'
'End Sub


Sub MakeHighlights_Click(control As IRibbonControl)

 ' Call ColourHighlights(False)
   Call Unavailable("")
 
End Sub

Sub ClearHighlights_Click(control As IRibbonControl)

'  Call ColourHighlights(True)
   Call Unavailable("")


End Sub


'====


Sub ColourHighlights(doClear As Boolean)

 ' On Error Resume Next
    
   ' result = MsgBox("This function will remove all current data within the Workbook." & vbCrLf & vbCrLf & "This cannot be undone." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbInformation, ProgramName$)
    
'    Dim myRange As Range
'
'            For Each sh In Sheets
'                For Each subName In sh.Names
'
'            ' this is the bit that colours it in'
'
'                   If Left(subName.Name, 1) = "'" Then realRangeName = Replace(subName.Name, "'" & sh.Name & "'!", "")
'                   If Left(subName.Name, 1) <> "'" Then realRangeName = Replace(subName.Name, sh.Name & "!", "")
'
'                   On Error Resume Next
'                   Set myRange = sh.Range(realRangeName)
'
'                      If doClear = True Then
'                        GoSub ClearColour
'
'
'
'                 ElseIf LCase(Left(realRangeName, 5)) = "t4pm_" And LCase(Mid(realRangeName, 8, 2)) = "r_" And myRange.Interior.Color = RGB(0, 255, 0) Then
'
'                      myRange = sh.Range(subName.Name)
'                       myColour1 = RGB(0, 255, 0)
'                      myColour2 = RGB(255, 0, 0)
'                      GoSub MakeStripes
'
'                 ElseIf LCase(Left(realRangeName, 5)) = "t4pm_" And LCase(Mid(realRangeName, 8, 2)) = "w_" And myRange.Interior.Color = RGB(255, 0, 0) Then
'                       myRange = sh.Range(subName.Name)
'                       myColour1 = RGB(0, 255, 0)
'                       myColour2 = RGB(255, 0, 0)
'                      GoSub MakeStripes
'
'
'                  ElseIf LCase(Left(realRangeName, 5)) = "t4pm_" And LCase(Mid(realRangeName, 8, 2)) = "r_" _
'                        And myRange.Interior.Pattern = xlNone Then
'
'                        myColour1 = RGB(255, 0, 0)
'                        GoSub MakeColour
'
'                  ElseIf LCase(Left(realRangeName, 5)) = "t4pm_" And LCase(Mid(realRangeName, 8, 2)) = "w_" _
'                        And myRange.Interior.Pattern = xlNone Then
'
'                        myColour1 = RGB(0, 255, 0)
'                        GoSub MakeColour
'
'
'                  End If
'
'                Next
'            Next
'
'Exit Sub'
'
'MakeColour:
'
'    With myRange.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = myColour1
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
 '   End With
'
'Return

'MakeStripes:
'    With myRange.Interior
'        .Pattern = xlDown
'        .PatternColor = myColour2
'        .Color = myColour1
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
''    End With
'
'
'Return
    
    
ClearColour:

    With myRange.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
        
    End With
    
Return

End Sub
