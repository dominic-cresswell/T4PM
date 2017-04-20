VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FieldRefForm 
   Caption         =   "Select Field Reference"
   ClientHeight    =   5055
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   5970
   OleObjectBlob   =   "FieldRefForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FieldRefForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CollectionDataOption_Click()
SingleDataOption = False
CollectionDataOption = True

Call RefreshFieldList("")
Call OptionsBoxUpdate("")

End Sub

Private Sub CreateField_Click()


    FieldRefOutput$ = ""
    
  ' intro
    FieldRefOutput$ = FieldRefOutput$ & "T4PM_"
    
  ' collection / single
    If FieldRefForm.CollectionDataOption = True Then
        FieldRefOutput$ = FieldRefOutput$ & "C_"
    Else
        FieldRefOutput$ = FieldRefOutput$ & "S_"
    End If
    
  ' read/write
    If FieldRefForm.DownloadCheck = True Then
    FieldRefOutput$ = FieldRefOutput$ & "R_"
    Else
    FieldRefOutput$ = FieldRefOutput$ & "W_"
    End If
    
  ' name!!
  
    GetListData$ = ""
    On Error Resume Next
    GetListData$ = CStr(FieldRefForm.FieldRefList.Value)

    If GetListData$ = "" Then
        result = MsgBox("No Field Data Selected", vbCritical, ProgramName$)
        Exit Sub
    End If
    
    GetListData$ = ClearSpecialCharacters(GetListData$)
          
    FieldRefOutput$ = FieldRefOutput$ & GetListData$ & "_"
    


   ' special rules
    Appendage$ = ""
    
    ' check for 'option'
    Select Case LCase(OptionsBox.Value)
    
    Case ""
        If CStr(NumberCombo.Value) = "" Then
        Appendage$ = "Null"
        Else
        End If
        
    Case "none"
        If CStr(NumberCombo.Value) = "" Or NumberCombo.Value = "n/a" Then
        Appendage$ = "Null"
        Else
        End If
    
    Case "grouped with line-ends"
        Appendage$ = "fLine"
    
    Case "grouped on single line"
        Appendage$ = "fSingle"

    Case "all items as list"
        Appendage$ = "nList"
        
    Case "number of items"
        Appendage$ = "nCount"
                
    Case "average of values"
        Appendage$ = "nAverage"
        
    Case "highest value"
        Appendage$ = "nHighest"
        
    Case "lowest value"
        Appendage$ = "nLowest"
        
    Case "sum of values"
        Appendage$ = "nSum"
        
    Case "earliest date"
        Appendage$ = "nFirst"
        
    Case "latest date"
        Appendage$ = "nLast"
        
    Case "date period (days)"
        Appendage$ = "nDaysLength"
        
    Case "date period (weeks)"
        Appendage$ = "nWeeksLength"

    Case Else
        Appendage$ = "Unknown"
        
    End Select
 
 ' check for 'item number'
    If NumberCombo.Value = "" Or NumberCombo.Value = "n/a" Then
        If Appendage$ = "" Then Appendage$ = "Null"
    Else
    
        GetData$ = NumberCombo.Value
        RemainderData$ = Replace(LCase(GetData$), "item", "")
        RemainderData$ = Replace(LCase(RemainderData$), ":", "")
        RemainderData$ = Replace(LCase(RemainderData$), " ", "")
        RemainderData$ = Replace(LCase(RemainderData$), "e", "")
        
        If IsNumeric(RemainderData$) = True Then
            Dim tempLng As Long
            tempLng = CLng(RemainderData$)
        
            If tempLng <> 0 Then
            Appendage$ = "n" & CStr(tempLng)
            Else
            Appendage$ = "Null"
            End If
        End If
        
    End If
    
    
    FieldRefOutput$ = FieldRefOutput$ & Appendage$
 
    If FieldRefOutput$ <> "" Then FieldRefForm.Hide
    

End Sub

Private Sub FieldRefList_Click()
    Call TextFieldUpdate("")
    Call OptionsBoxUpdate("")

End Sub



Private Sub NumberCombo_AfterUpdate()

    If MultipleBox.Value = False Then Exit Sub
    
    GetData$ = NumberCombo.Value
    
    RemainderData$ = Replace(LCase(GetData$), "item", "")
    RemainderData$ = Replace(LCase(RemainderData$), ":", "")
    RemainderData$ = Replace(LCase(RemainderData$), " ", "")
    RemainderData$ = Replace(LCase(RemainderData$), "e", "")
    
  On Error Resume Next
  
    'standard
    If Left(LCase(GetData$), 5) = "item:" And IsNumeric(RemainderData$) = True Then
        If CLng(RemainderData$) < 1 Then
            NumberCombo.Value = "Item: 001"
        ElseIf CLng(RemainderData$) < 1000 Then
            NumberCombo.Value = "Item: " & Format(CLng(RemainderData$), "000")
        End If
        
        Exit Sub
    End If
    
    ' other
    
    
    If IsNumeric(GetData$) = True Then
        NumberCombo.Value = Replace(LCase(NumberCombo.Value), "e", "")
        
        NumberCombo.Value = "Item: " & Format(CLng(GetData$), "#,###,000")
        Exit Sub
    End If
    
    result = MsgBox("Enter numeric values only", vbCritical, ProgramName$)
    
    
    
    
    
    
    
    
    
End Sub

Private Sub OptionsBox_Change()
 
 If FieldRefForm.MultipleBox = True Then
 
    ' no combo option (selected option)
    If FieldRefForm.OptionsBox <> "None" Then
            FieldRefForm.NumberCombo = ""
            FieldRefForm.NumberCombo.Enabled = False
            
    Else
            FieldRefForm.NumberCombo.Clear

                FieldRefForm.NumberCombo.Enabled = True
                FieldRefForm.NumberCombo.MatchEntry = fmMatchEntryNone
            
            For nn = 1 To 256
                FieldRefForm.NumberCombo.AddItem "Item: " & Format(nn, "000")
            Next
            
            FieldRefForm.NumberCombo = "Item: 001"
    End If
    
Else        ' no combo option
            FieldRefForm.NumberCombo = ""
            FieldRefForm.NumberCombo.Enabled = False
    
 End If

End Sub

Private Sub SingleDataOption_Click()

SingleDataOption = True
CollectionDataOption = False

Call RefreshFieldList("")
Call OptionsBoxUpdate("")

End Sub

Private Sub UserForm_Initialize()

SingleDataOption = True
CollectionDataOption = False

Call RefreshFieldList("")
Call OptionsBoxUpdate("")
End Sub
