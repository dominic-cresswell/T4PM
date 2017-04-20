Attribute VB_Name = "PMToolKitCode3"

Public ProjectWriteDataArray(9999, 4)
Public ProjectReadDataArray(9999, 4)


Sub PullWriteDataFromWorksheets(dummy$)

' get the valid field data
    Call ImportFieldData("")

On Error Resume Next
 '   For Each sh In Sheets
    For Each sh In Excel.ActiveWindow.SelectedSheets
        '  MsgBox sh.Name
        '  Stop
        '  If sh.Select = False Then GoTo skip:
          
        For Each bitname In sh.Names
            FoundRef$ = ""
            FoundRef$ = bitname.Name
            
            
            ' check it is a T4PM range name
            If InStr(vbTextCompare, FoundRef$, "T4PM") > 0 Then
                
               sheetlength = Len(sh.Name)
               If Left(FoundRef$, 1) = "'" And Mid(FoundRef$, sheetlength + 2, 1) = "'" Then sheetlength = sheetlength + 2
               
               FoundRef$ = Right(FoundRef$, Len(FoundRef$) - sheetlength)
               FoundRef$ = Replace(FoundRef$, "!T4PM_", "")
               
            ' check it is  single bit of data
         If LCase(Left(FoundRef$, 2)) = "s_" Then
               FoundRef$ = Right(FoundRef$, Len(FoundRef$) - 2)

          ' check it is  writable bit of data
          If LCase(Left(FoundRef$, 2)) = "w_" Then
                FoundRef$ = Right(FoundRef$, Len(FoundRef$) - 2)

            ' if it ends 'null' make it _n0
             If LCase(Right(FoundRef$, 5)) = "_null" Then
                FoundRef$ = Left(FoundRef$, Len(FoundRef$) - 5) & "_n0"
             End If
                                     
              ' all checks passed
               GoSub PlaceInArray:
               
          Else    ' not writable
            'Exit For
          End If
         
         Else    ' not a single item
            'Exit For
         End If
               
        End If
        Next
        
skip:
    Next


' =====  check we have a permitted user

    For nnn = 0 To 9999
        TempField$ = LCase(ProjectWriteDataArray(nnn, 0))
        
            If Left(LCase(ProjectWriteDataArray(nnn, 0)), Len("permittedusers")) = "permittedusers" Then
                  Exit For
            ElseIf ProjectWriteDataArray(nnn, 0) = "" Then
                ProjectWriteDataArray(nnn, 0) = "PermittedUsers_n1"
                ProjectWriteDataArray(nnn, 1) = Environ("username")
                ProjectWriteDataArray(nnn, 2) = "text"
                Exit For
            
            End If
        
        
    Next nnn
    

Exit Sub


' we will store only:  Valid Data, Single Items, Writable Items (TO DO)

PlaceInArray:

    For nnn = 0 To 9999
    
        ' this is the check for finding data already placed in the array
        If LCase(ProjectWriteDataArray(nnn, 0)) = LCase(FoundRef$) Then
            TempRef$ = Left(FoundRef$, InStr(vbTextCompare, FoundRef$, "_") - 1)
            
            ' get name and worksheet seperately
            GetRangeName$ = ""
            GetRangeSheet$ = ""
            GetRangeName$ = Right(bitname.Name, Len(bitname.Name) - InStr(vbTextCompare, bitname.Name, "!"))
            GetRangeSheet$ = Left(bitname.Name, InStr(vbTextCompare, bitname.Name, "!") - 1)
            
            If Left(GetRangeSheet$, 1) = "'" And Right(GetRangeSheet$, 1) = "'" Then
            GetRangeSheet$ = Mid(GetRangeSheet$, 2, Len(GetRangeSheet$) - 2)
            End If
            
            ' do the checks!
             If ProjectWriteDataArray(nnn, 0) <> "" _
                And ProjectWriteDataArray(nnn, 1) <> Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells(1).Text _
                And Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells(1).Text <> "" Then
 '               And Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells(1).Text Is Not Null Then
                
                longmsg$ = ""
                longmsg$ = longmsg$ & "Beware: Repeated data (" & TempRef$ & ") is entered with varying values "
                longmsg$ = longmsg$ & "and has not been uploaded when repeated." & vbCrLf & vbCrLf
                longmsg$ = longmsg$ & "Stored value: '" & ProjectWriteDataArray(nnn, 1) & "'" & vbCrLf
                longmsg$ = longmsg$ & "Second value: '" & Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells(1).Text & "'  " & vbCrLf
                longmsg$ = longmsg$ & "Worksheet: '" & GetRangeSheet$ & "'" & vbCrLf
                 longmsg$ = longmsg$ & "Cell Name: '" & GetRangeName$ & "' "
                
                result = MsgBox(longmsg$, vbCritical, ProgramName$)
            End If
            
            ' at this point, we passed, but we dont care.
        
        
        ' this is for 'new' data
        ElseIf ProjectWriteDataArray(nnn, 0) = "" Or LCase(ProjectWriteDataArray(nnn, 0)) = LCase(FoundRef$) Then
        
        TempRef$ = Left(FoundRef$, InStr(vbTextCompare, FoundRef$, "_") - 1)
        'Len(TempRef$) -
        
        ' check if the data differs
            If ProjectWriteDataArray(nnn, 0) <> "" And ProjectWriteDataArray(nnn, 0) <> FoundRef$ Then
               result = MsgBox("Beware: Repeated data (" & TempRef$ & ") is entered with varying values.", vbCritical, ProgramName$)
            End If
        
    
        
        ProjectWriteDataArray(nnn, 0) = FoundRef$
        ProjectWriteDataArray(nnn, 1) = ""
        
        ' get name and worksheet seperately
            GetRangeName$ = ""
            GetRangeSheet$ = ""
            GetRangeName$ = Right(bitname.Name, Len(bitname.Name) - InStr(vbTextCompare, bitname.Name, "!"))
            GetRangeSheet$ = Left(bitname.Name, InStr(vbTextCompare, bitname.Name, "!") - 1)
            
            If Left(GetRangeSheet$, 1) = "'" And Right(GetRangeSheet$, 1) = "'" Then
            GetRangeSheet$ = Mid(GetRangeSheet$, 2, Len(GetRangeSheet$) - 2)
            End If
            
            
            ' this is a fix for merged cells
            If Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells.Count > 1 Then
            ProjectWriteDataArray(nnn, 1) = Worksheets(GetRangeSheet$).Range(GetRangeName$).Cells(1).Text
            Else
            ProjectWriteDataArray(nnn, 1) = Worksheets(GetRangeSheet$).Range(GetRangeName$).Text
            End If


        ' validate the datatype against inputted info
         ProjectWriteDataArray(nnn, 4) = ""
         GoSub ValidateData:
        
        
      ' check we are not just recording a 'blank'
        If IsNull(ProjectWriteDataArray(nnn, 1)) = True Or ProjectWriteDataArray(nnn, 1) = Empty Then
            If ProjectWriteDataArray(nnn, 4) = "" Then ProjectWriteDataArray(nnn, 4) = "error-null"
        End If
        
        ' plonk the data down
            If ProjectWriteDataArray(nnn, 4) = "error-validate" Then
                 
                 ProjectWriteDataArray(nnn, 0) = ""
                 ProjectWriteDataArray(nnn, 1) = ""
                 ProjectWriteDataArray(nnn, 2) = ""
                 ProjectWriteDataArray(nnn, 3) = ""
                 ProjectWriteDataArray(nnn, 4) = ""
                 
                result = MsgBox("Beware: Data for (" & TempRef$ & ") does not match data type and will not be stored.", vbCritical, ProgramName$)
                Exit For
            ElseIf ProjectWriteDataArray(nnn, 4) = "error-null" Then
                 ProjectWriteDataArray(nnn, 0) = ""
                 ProjectWriteDataArray(nnn, 1) = ""
                 ProjectWriteDataArray(nnn, 2) = ""
                 ProjectWriteDataArray(nnn, 3) = ""
                 ProjectWriteDataArray(nnn, 4) = ""
                 
               ' Result = MsgBox("Beware: Data for (" & TempRef$ & ") does not match data type and will not be stored.", vbCritical, ProgramName$)
                Exit For
            End If
        
        Exit For
        End If
    Next

Return
    
ValidateData:

       ' TempRef$
         For zzz = 0 To 9999
            
            TempField$ = LCase(FieldListArray(zzz, 1))
            TempField$ = ClearSpecialCharacters(TempField$)
         If TempField$ = "" Then Exit For
            
            If TempField$ = Left(LCase(FoundRef$), Len(TempField$)) Then
            
        ' Debug.Print TempField$; "   =   "; FieldListArray(zzz, 2)
                ProjectWriteDataArray(nnn, 2) = FieldListArray(zzz, 2)
                
            
                ' now check that is conforms
                Select Case ProjectWriteDataArray(nnn, 2)
                
                Case "text"
                
                
                Case "memo"
                
                
                Case "date"
                    If IsDate(ProjectWriteDataArray(nnn, 1)) = False And ProjectWriteDataArray(nnn, 1) <> "" Then
                       ProjectWriteDataArray(nnn, 4) = "error-validate"
                    End If
                    
                    ProjectWriteDataArray(nnn, 1) = Format(ProjectWriteDataArray(nnn, 1), "dd-mmmm-yyyy")
                    
                Case "numeric"
                    If IsNumeric(ProjectWriteDataArray(nnn, 1)) = False Then
                       ProjectWriteDataArray(nnn, 4) = "error-validate"
                    End If
                
                Case "currency"
                
                     If IsNumeric(ProjectWriteDataArray(nnn, 1)) = False Then
                       ProjectWriteDataArray(nnn, 4) = "error-validate"
                    End If
                    
                    ProjectWriteDataArray(nnn, 1) = Format(ProjectWriteDataArray(nnn, 1), "£#,###,###,##0.00")
                
                Case "boolean"
                
                If LCase(ProjectWriteDataArray(nnn, 1)) = "yes" _
                Or LCase(ProjectWriteDataArray(nnn, 1)) = "y" _
                Or LCase(ProjectWriteDataArray(nnn, 1)) = "true" Then
                    ProjectWriteDataArray(nnn, 1) = True
                    
            ElseIf LCase(ProjectWriteDataArray(nnn, 1)) = "no" _
                Or LCase(ProjectWriteDataArray(nnn, 1)) = "n" _
                Or LCase(ProjectWriteDataArray(nnn, 1)) = "false" Then
                    ProjectWriteDataArray(nnn, 1) = False
                End If
                
                   If WorksheetFunction.IsLogical(ProjectWriteDataArray(nnn, 1)) = False Then
                       ProjectWriteDataArray(nnn, 4) = "error-validate"
                    End If
                
                End Select
                
                Exit For
            End If
         
         Next
         

Return
    
    
    
End Sub
            
         ''   ' check it is  single bit of data
         ''   If LCase(Left(FoundRef$, 2)) = "s_" Then
         ''   FoundRef$ = Right(FoundRef$, Len(FoundRef$) - 2)
          ''
           ''     If LCase(Left(FoundRef$, 2)) = "s_" Then
          ''      FoundRef$ = Right(FoundRef$, Len(FoundRef$) - 2)
       ''     Else
       ''     ' not a single item
       ''     End If
            
               
            
Function GetTempData(inData$) As String

    CheckField$ = inData$
    CheckField$ = LCase(CheckField$)
    CheckField$ = ClearSpecialCharacters(CheckField$)

    ' check the
         For zzz = 0 To 9999
                  
            TempField$ = LCase(ProjectWriteDataArray(zzz, 0))
            TempField$ = ClearSpecialCharacters(TempField$)
            
            
            If CheckField$ = Left(LCase(TempField$), Len(CheckField$)) Then
             '   Debug.Print "   =   "; ProjectWriteDataArray(zzz, 1)
                GetTempData = ProjectWriteDataArray(zzz, 1)
                Exit For
            End If
        
        Next

End Function

Function GetTempData2(inData$) As String

    CheckField$ = inData$
    CheckField$ = LCase(CheckField$)
    CheckField$ = ClearSpecialCharacters(CheckField$)

    ' check the
         For zzz = 0 To 9999
                  
            TempField$ = LCase(ProjectReadDataArray(zzz, 0))
            TempField$ = ClearSpecialCharacters(TempField$)
            
            
            If CheckField$ = Left(LCase(TempField$), Len(CheckField$)) Then
             '   Debug.Print "   =   "; ProjectWriteDataArray(zzz, 1)
                GetTempData2 = ProjectReadDataArray(zzz, 1)
                Exit For
            End If
        
        Next

End Function


