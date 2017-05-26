Attribute VB_Name = "E_DataUploadCode"
Option Private Module

Function GetLiveReferenceCode() As String

    For Each sh In ActiveWorkbook.Sheets
        For Each nm In sh.Names
        
            If InStr(vbTextCompare, LCase(nm.Name), "t4pm") > 0 And InStr(vbTextCompare, LCase(nm.Name), "projectreference") Then
            
            sheetname$ = sh.Name
            rangename$ = nm.Name
            rangename$ = Replace(rangename$, sheetname$, "")
            rangename$ = Replace(rangename$, "''", "")
            If Left(rangename$, 1) = "!" Then rangename$ = Right(rangename$, Len(rangename$) - 1)
            
                On Error Resume Next
                oldrefcode = GetLiveReferenceCode
                GetLiveReferenceCode = sh.Range(rangename$).Cells(1).Value
                If oldrefcode <> GetLiveReferenceCode Then refCount = refCount + 1

            End If
            
        Next
    
    Next


    If refCount > 1 Then
        GetLiveReferenceCode = ""
        Result = MsgBox("There are multiple Reference Codes in the Active Workbook.", vbCritical, ProgramName$)
    End If
    

End Function



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



Sub ExportDataToStore(showmsg As Boolean)

    ' ======
    If CurrentStore$ = "" Or FileExists(CurrentStore$) = False Then
        Result = MsgBox("No T4PM Project Store selected.", vbCritical, ProgramName$)
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
