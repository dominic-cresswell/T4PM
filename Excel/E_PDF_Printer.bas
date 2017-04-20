Attribute VB_Name = "E_PDF_Printer"




Sub MakePDFSheet(dummy$)
    
   GetSheetName$ = ClearSpecialCharacters(ActiveSheet.Name)

    'MsgBox GetSheetName$


    On Error Resume Next

    IssueDateField$ = "T4PM_S_W_" & GetSheetName$ & "IssueDate_Null"
    RibaStageField$ = "T4PM_S_W_CurrentRibaStage_Null"

    On Error GoTo abort:

    Set IssueDateRange = ActiveSheet.Range(IssueDateField$)
    Set RibaStageRange = ActiveSheet.Range(RibaStageField$)


    If RibaStageRange.Text = "" Or IsNumeric(RibaStageRange.Text) = False Then
        result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
        Exit Sub

    End If

    
    If CLng(RibaStageRange.Text) < 0 Or CLng(RibaStageRange.Text) > 7 Then

        result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
        Exit Sub

    End If
    

   '  If IssueDateRange.Text = "" Or IssueDateRange(RibaStageRange.Text) = False Then
   '     result = MsgBox("Invalid RIBA stage Number.", vbCritical, ProgramName$)
    '    Exit Sub
   ' End If

   On Error GoTo 0

    IssueDateRange.Value = Format(Date, "dd-mm-yyyy")
    SaveFieldDate$ = Format(Date, "dd-mm-yyyy")
    SaveFieldName$ = GetSheetName$ & "_Stage" & RibaStageRange.Text & "_n0"

    result = MsgBox("Data will been stored, stating this has been issued as: " & vbCrLf & SaveFieldName$ & " - " & SaveFieldDate$, vbInformation, ProgramName$)
    
    ' Save to PDF?

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        AddSlash(ActiveWorkbook.Path) & "" & Replace(SaveFieldName$, "_n0", "") & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True


    ' store the data to the store....


    GoSub StolenCode:


Exit Sub

abort:

        result = MsgBox("This template is not valid for PDF issuing." & vbCrLf & "No RIBA Stage or Issue Date field on this Worksheet.", vbCritical, ProgramName$)

        Exit Sub

        

        

StolenCode:

' invoke a new Excel

    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")

    exlApp.visible = False

    If CurrentStore$ = "" Then
            result = MsgBox("No Project Store selected.", vbCritical, ProgramName$)
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

        For qqq = 1 To 9999

             If exlSheet.Columns(1).Rows(qqq) = "" Or LCase(exlSheet.Columns(1).Rows(qqq)) = LCase(SaveFieldName$) Then
                exlSheet.Columns(1).Rows(qqq) = SaveFieldName$
                exlSheet.Columns(2).Rows(qqq) = SaveFieldDate$
                exlSheet.Columns(3).Rows(qqq) = Format(Date + Time, "dd-mmm-yyyy hh:mm")
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
    result = MsgBox("No worksheet 'Project Store' within working store.", vbCritical, ProgramName$)

End Sub

