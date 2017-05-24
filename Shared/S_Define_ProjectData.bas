Attribute VB_Name = "S_Define_ProjectData"
Option Private Module

Public ProjectWriteDataArray(9999, 4)
Public ProjectReadDataArray(9999, 4)


' Define user-defined type.
Public Type ProjectData
    SiteName As String
    ProjectDescription As String
    ProjectReference As String
    AllUsers As String
End Type


' =============== pull key data from File.


Function GetTopData(inFile$) As ProjectData
    Dim prjd As ProjectData
    
    prjd.SiteName = ""
    prjd.ProjectDescription = ""
    prjd.ProjectReference = ""
    prjd.AllUsers = ""

' invoke a new Excel
    Dim exlApp As Excel.Application
    Set exlApp = CreateObject("Excel.Application")
    exlApp.visible = False
    
    If inFile$ = "" Then Exit Function
    
       
    Dim exlDoc As Workbook
    On Error GoTo fail2:
    Set exlDoc = exlApp.Workbooks.Open(inFile$)

    
    Dim exlSheet As Worksheet
    On Error Resume Next
    Set exlSheet = exlDoc.Worksheets.Item("ProjectStore")
    
    If exlSheet Is Nothing Then GoTo fail:


   ' we made it!

      ' find existing data
        For QQQ = 1 To 9999

                FieldName$ = exlSheet.Columns(1).Rows(QQQ)
                FieldData$ = exlSheet.Columns(2).Rows(QQQ)
            '    FieldStamp$ = exlSheet.Columns(3).Rows(qqq)
                
                If FieldName$ = "" Then Exit For
                
             '   ProjectReadDataArray(zzz, 0) = FieldName$
             '   ProjectReadDataArray(zzz, 1) = FieldData$
               ' ProjectReadDataArray(zzz, 2) = FieldStamp$
                
                
                If FieldName$ = "SiteName_n0" Then prjd.SiteName = FieldData$
                If FieldName$ = "ProjectDescription_n0" Then prjd.ProjectDescription = FieldData$
                If FieldName$ = "ProjectReference_n0" Then prjd.ProjectReference = FieldData$
                
                If Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers = "" Then
                    prjd.AllUsers = FieldData$
                ElseIf Left(FieldName$, 16) = "PermittedUsers_n" And prjd.AllUsers <> "" Then
                    prjd.AllUsers = prjd.AllUsers & ", " & FieldData$
                End If

                zzz = zzz + 1
        Next

    exlDoc.Close (False)
    exlApp.Quit
    
    
    GetTopData = prjd
    
    Exit Function
    

fail:
    exlDoc.Close
fail2:
    exlApp.Quit
  
        
    GetTopData = prjd
    
End Function
