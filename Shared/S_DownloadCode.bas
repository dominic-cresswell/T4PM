Attribute VB_Name = "S_DownloadCode"

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
  


