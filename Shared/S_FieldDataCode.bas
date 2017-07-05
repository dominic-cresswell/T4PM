Attribute VB_Name = "S_FieldDataCode"
Public FieldRefOutput$




Sub SetProjectFolder(dummy$)


   ' cehck we have a store selected, if not, grab it again
        If CurrentStore$ = "" Then
            Call RestoreStore("")
        End If
        
          If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
               Result = MsgBox("Please re-select T4PM Project Store", vbCritical, ProgramName$)
               Exit Sub
        End If
                
        FindFolder$ = GetFolder("G:\CED PG\Premises\")
        
        If FindFolder$ = "" Or DirExists(FindFolder$) = False Then
              Result = MsgBox("Invalid Folder selection.", vbCritical, ProgramName$)
              Exit Sub
        End If
    
        FindFolder$ = AddSlash(FindFolder$)
        
        Call ClearWriteData("")

        ProjectWriteDataArray(nnn, 0) = "FolderPath_n0"   ' field name
        ProjectWriteDataArray(nnn, 1) = FindFolder$     ' content / value
        ProjectWriteDataArray(nnn, 2) = "text"          ' type
        ProjectWriteDataArray(nnn, 3) = ""
        ProjectWriteDataArray(nnn, 4) = ""              ' errors
        
        Call ExportDataToStore(True)
                
End Sub



Sub OptionsBoxUpdate(dummy$)

    FieldRefForm.OptionsBox.Clear


    Select Case LCase(FieldRefForm.DataTypeText)
    Case "collection"
        
        FieldRefForm.OptionsBox.AddItem "None"
        FieldRefForm.OptionsBox.AddItem "Grouped with Line-ends"
        FieldRefForm.OptionsBox.AddItem "Grouped on Single Line"
    
        GoSub nullNumber:

    
    Case "text"
    
        FieldRefForm.OptionsBox.AddItem "None"
        
        If FieldRefForm.MultipleBox = True Then
            FieldRefForm.OptionsBox.AddItem "All Items as List"
            
            GoSub ListNumber:

        Else
            GoSub nullNumber:

        End If
        
        
        
    Case "date"
    
        FieldRefForm.OptionsBox.AddItem "None"
        If FieldRefForm.MultipleBox = True Then
            
            FieldRefForm.OptionsBox.AddItem "All Items as List"
            FieldRefForm.OptionsBox.AddItem "Earliest Date"
            FieldRefForm.OptionsBox.AddItem "Latest Date"
            FieldRefForm.OptionsBox.AddItem "Date Period (days)"
            FieldRefForm.OptionsBox.AddItem "Date Period (weeks)"
            
            GoSub ListNumber:
 
        Else
            GoSub nullNumber:

        End If
        
    Case "numeric"
    
        FieldRefForm.OptionsBox.AddItem "None"
        If FieldRefForm.MultipleBox = True Then
            
            FieldRefForm.OptionsBox.AddItem "All Items as List"
            FieldRefForm.OptionsBox.AddItem "Number of Items"
            FieldRefForm.OptionsBox.AddItem "Average of Values"
            FieldRefForm.OptionsBox.AddItem "Highest Value"
            FieldRefForm.OptionsBox.AddItem "Lowest Value"
            FieldRefForm.OptionsBox.AddItem "Sum of Values"

            GoSub ListNumber:

            
        Else
            GoSub nullNumber:

        End If
        
    Case "Currency"
    
        FieldRefForm.OptionsBox.AddItem "None"
        If FieldRefForm.MultipleBox = True Then
            
            FieldRefForm.OptionsBox.AddItem "All Items as List"
            FieldRefForm.OptionsBox.AddItem "Number of Items"
            FieldRefForm.OptionsBox.AddItem "Highest Value"
            FieldRefForm.OptionsBox.AddItem "Lowest Value"
            FieldRefForm.OptionsBox.AddItem "Sum of Values"
            
            GoSub ListNumber:
            
        Else
            GoSub nullNumber:
        End If
        
    End Select
    
    FieldRefForm.OptionsBox = "None"
    
Exit Sub


nullNumber:
            FieldRefForm.NumberCombo.Clear
            FieldRefForm.NumberCombo.AddItem "n/a"
            FieldRefForm.NumberCombo.Value = "n/a"
            FieldRefForm.NumberCombo.Enabled = False
                        
Return

ListNumber:
            FieldRefForm.NumberCombo.Enabled = True
            FieldRefForm.NumberCombo.MatchEntry = fmMatchEntryNone
            
            FieldRefForm.NumberCombo.Clear
            For nn = 1 To 256
                FieldRefForm.NumberCombo.AddItem "Item: " & Format(nn, "000")
            Next
            
            FieldRefForm.NumberCombo = "Item: 001"
            
Return

End Sub

Sub TextFieldUpdate(dummy$)

    If FieldRefForm.FieldRefList.ListIndex = 0 And FieldRefForm.FieldRefList.Selected(0) = False Then
        FieldRefForm.DataTypeText = "No field Selected"
        FieldRefForm.RefText = "n/a"
        FieldRefForm.MultipleBox = False
        Exit Sub
    End If
    

    If FieldRefForm.CollectionDataOption = True Then
        FieldRefForm.DataTypeText = "Collection"
        FieldRefForm.RefText = "n/a"
        FieldRefForm.MultipleBox = False
        Exit Sub
    End If
    
    If FieldRefForm.SingleDataOption = True Then
            Z = FieldRefForm.FieldRefList.ListIndex
            
            FieldRefForm.DataTypeText = FieldRefForm.FieldRefList.Column(1, Z)
            FieldRefForm.RefText = FieldRefForm.FieldRefList.Column(2, Z)
            FieldRefForm.MultipleBox = FieldRefForm.FieldRefList.Column(3, Z)
    
    End If
    
    
End Sub



Sub GetCollectionList(dummy$)


    If FieldListArray(0, 1) = "" Then Call ImportFieldData("")
        
            n = 0
            Do Until FieldListArray(n, 1) = ""
                If FieldListArray(n, 3) <> "" Then
                    GoSub PushCollectionItem
                End If
                
                    n = n + 1
            Loop

            
Exit Sub

PushCollectionItem:
    
        For nn = 0 To 9999
        If CollectionListArray(nn) = "" Or LCase(CollectionListArray(nn)) = LCase(FieldListArray(n, 3)) Then
                CollectionListArray(nn) = FieldListArray(n, 3)
                Exit For
        End If
        Next
    
Return

End Sub

Sub RefreshFieldList(dummy$)

    FieldRefForm.FieldRefList.Clear
    
        If FieldRefForm.SingleDataOption = True Then
        
                n = 0
                
            Do Until FieldListArray(n, 1) = ""
                FieldRefForm.FieldRefList.AddItem (FieldListArray(n, 1))
                Z = FieldRefForm.FieldRefList.ListCount - 1
                
                FieldRefForm.FieldRefList.Column(1, Z) = FieldListArray(n, 2)
                FieldRefForm.FieldRefList.Column(2, Z) = FieldListArray(n, 0)
                FieldRefForm.FieldRefList.Column(3, Z) = FieldListArray(n, 4)
                
                n = n + 1
            Loop
            
        ElseIf FieldRefForm.CollectionDataOption = True Then
        
            Call GetCollectionList("")
            
            Do Until CollectionListArray(n) = ""
                FieldRefForm.FieldRefList.AddItem (CollectionListArray(n))
                n = n + 1
            Loop
            
            
        End If


End Sub



Sub ImportFieldData(dummy$)

   ' but we must clear the list first!!
    For rrr = 0 To 9999
        For eee = 0 To 4
            FieldListArray(rrr, eee) = ""
        Next
    Next

    MajorFieldData$ = ""
  
    ' routine check
    ProgramPath$ = S_UserConfigCode.CheckProgramPath
        
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

