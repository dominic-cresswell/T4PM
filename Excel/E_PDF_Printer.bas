Attribute VB_Name = "E_PDF_Printer"




Sub MakePDFSheet(dummy$)
    
   GetSheetName$ = ClearSpecialCharacters(ActiveSheet.Name)


'====
    On Error Resume Next

    IssueDateField$ = "T4PM_S_W_" & GetSheetName$ & "IssueDate_Null"
    Append$ = "IssueDate"
    
    Set IssueDateRange = ActiveSheet.Range(IssueDateField$)
    
    If IssueDateRange Is Nothing Then
       IssueDateField$ = "T4PM_S_W_" & GetSheetName$ & "FormUpdated_Null"
        Append$ = "FormUpdated"
        Set IssueDateRange = ActiveSheet.Range(IssueDateField$)
    End If
                     
'==
    RibaStageField$ = "T4PM_S_W_CurrentRibaStage_Null"
    Set RibaStageRange = ActiveSheet.Range(RibaStageField$)


'+++++++++++++
    
    If IssueDateRange Is Nothing Then GoTo abort:
    
    If RibaStageRange Is Nothing Then
    Else
    
        If RibaStageRange.Text = "" Or IsNumeric(RibaStageRange.Text) = False Then
            Result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
            Exit Sub
        End If

        If CLng(RibaStageRange.Text) < 0 Or CLng(RibaStageRange.Text) > 7 Then
            Result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
            Exit Sub
        End If
    End If


   '  If IssueDateRange.Text = "" Or IssueDateRange(RibaStageRange.Text) = False Then
   '     result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
    '    Exit Sub
   ' End If

   On Error GoTo 0

    IssueDateRange.Value = Format(Date, "dd-mm-yyyy")
    SaveFieldDate$ = Format(Date, "dd-mm-yyyy")
    SaveFieldName$ = GetSheetName$
    
    On Error Resume Next
    SaveFieldName$ = SaveFieldName$ & "_Stage" & RibaStageRange.Text
    ACount = 0
    

    Result = MsgBox("Data will been stored, stating this has been issued as: " & vbCrLf & SaveFieldName$ & "_n" & ACount & "_" & SaveFieldDate$, vbInformation, ProgramName$)
    
    ' Save to PDF?

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        AddSlash(ActiveWorkbook.Path) & "" & Replace(SaveFieldName$ & "_n" & ACount & "_" & SaveFieldDate$, "", "") & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True


    ' store the data to the store....


    GoSub StolenCode:


Exit Sub

abort:

        Result = MsgBox("This template is not valid for PDF issuing." & vbCrLf & "No Completed or Issue Date field on this Worksheet.", vbCritical, ProgramName$)

        Exit Sub

        

        

StolenCode:

' invoke a new Excel

    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")

    exlApp.visible = False

    If CurrentStore$ = "" Then
            Result = MsgBox("No Project Store selected.", vbCritical, ProgramName$)
        Exit Sub

    End If

       

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

             If exlSheet.Columns(1).Rows(QQQ) = "" Or LCase(exlSheet.Columns(1).Rows(QQQ)) = LCase(SaveFieldName$) Then
                exlSheet.Columns(1).Rows(QQQ) = SaveFieldName$ & Append$ & "_n" & ACount
                exlSheet.Columns(2).Rows(QQQ) = SaveFieldDate$
                exlSheet.Columns(3).Rows(QQQ) = Format(Date + Time, "dd-mmm-yyyy hh:mm")
                Exit For
              End If
     
        Next

    exlDoc.Save
    exlDoc.Close (True)
    exlApp.Quit

  '  If showmsg = True Then Result = MsgBox("Data Downloaded", vbInformation, ProgramName$)

    Exit Sub

fail:
    exlDoc.Close

fail2:

    exlApp.Quit
    Result = MsgBox("No worksheet 'Project Store' within working store.", vbCritical, ProgramName$)

End Sub



