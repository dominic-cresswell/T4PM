Attribute VB_Name = "S_DownloadCode"



Sub RestoreStore(dummy$)

' the path
    If ProgramPath$ = "" Then ProgramPath$ = AddSlash(S_UserConfigCode.CheckProgramPath)

   ' restore using RememberProject
'    If RememberProject = True Then
    If FileExists(ProgramPath$ & "LastProject") = True Then
        getProject$ = ReadTextFile(ProgramPath$ & "LastProject")
        Call LoadProjectStore(getProject$)
        Exit Sub
    End If
    
    If Application.Name = "Microsoft Word" Then Exit Sub




  ' this is a nasty bodge, specifcally for the Excel version

    If Application.Name = "Microsoft Excel" Then

        ' grab the working path
                    If WorkingPath$ = "" Then
                        WorkingPath$ = AddSlash(GetConfigSetting("WorkingPath"))
                    End If
                    
                    If WorkingPath$ <> "" And DirExists(WorkingPath$) = True Then
                        TempStore$ = WorkingPath$
                    End If
                    
        ' grab the reference number
                 If E_UploadCode.GetTempData_WriteBuffer("Project Reference") = "" Then
                     Call PullWriteDataFromWorksheets("")
                 End If
             
                 
                 TempRef$ = E_UploadCode.GetTempData_WriteBuffer("Project Reference")
                        
                 If TempRef$ <> "" And TempStore$ <> "" Then
                     TempStore$ = TempStore$ & TempRef$
                 End If
                 
                 
        ' check file exists
                If FileExists("T4PM" & TempStore$ & ".xls") = False Then
                        TempStore$ = ""
                End If
            
                If TempStore$ <> "" Then CurrentStore$ = "T4PM" & TempStore$ & ".xls"
                Call LoadProjectStore(CurrentStore$)

    End If
End Sub





Sub LoadProjectStore(inFile$)

    inFile$ = Replace(inFile$, vbCrLf, "")
    inFile$ = Replace(inFile$, vbCr, "")
    inFile$ = Replace(inFile$, vbLf, "")
    
' check the file is a T4PM
    If InStr(vbTextCompare, LCase(inFile$), "t4pm_") < 1 Then
        Result = MsgBox("No valid T4PM Project Store selected", vbCritical, ProgramName$)
        Exit Sub
    End If
  
  ' check it exist!
  ' =======
    If FileExists(inFile$) = False Then
        Result = MsgBox("Invalid T4PM Project Store Selection", vbCritical, ProgramName$)
        Exit Sub
    End If
    

    ' but are we permitted to use this

    If VerifyStoreUsers(inFile$) = False Then
        inFile$ = ""
        Result = MsgBox("You are not a permitted user for this T4PM Project Store.", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    DoEvents
    
        CurrentStore$ = inFile$
    ' we got through!
    
     Call ClearReadData("")
     DoEvents
  '  Call RestoreStore("")
     Call ImportDataFromStore("")
        
        
     Call RefreshRibbon("")
    
    
Exit Sub
abort:
    Result = MsgBox("Invalid T4PM Project Store Selection", vbCritical, ProgramName$)
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



Function GetTempData_ReadBuffer(inData$) As String

    CheckField$ = inData$
    CheckField$ = LCase(CheckField$)
    CheckField$ = ClearSpecialCharacters(CheckField$)

    ' check the
         For zzz = 0 To 9999
                  
            TempField$ = LCase(ProjectReadDataArray(zzz, 0))
            TempField$ = ClearSpecialCharacters(TempField$)
            
            If CheckField$ = Left(LCase(TempField$), Len(CheckField$)) Then
             '   Debug.Print "   =   "; ProjectWriteDataArray(zzz, 1)
                GetTempData_ReadBuffer = ProjectReadDataArray(zzz, 1)
                Exit For
            End If
        
        Next

End Function






Function GetAnyDataForHeaders(inString$) As String

If inString$ = "" Then Exit Function

        ' check the 'write' (updated) data
    '   GetAnyDataForHeaders = CStr(GetTempData(inString$))
       
       ' check the 'read' (stored) data
       If GetAnyDataForHeaders = "" Then
       If ProjectReadDataArray(0, 0) = "" Then Call ImportDataFromStore("")
       GetAnyDataForHeaders = GetTempData_ReadBuffer(inString$)
       End If
  
End Function
  


