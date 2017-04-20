Attribute VB_Name = "S_RibbonStateCode"


Public MyTag As String
Public VisibleState As Boolean


#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If


#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
        Dim objRibbon As Object
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
End Function



Sub GetVisible(control As IRibbonControl, ByRef visible)
    If MyTag = "show" Then
        visible = True
    Else
        If control.Tag Like MyTag Then
            visible = True
        Else
            visible = False
        End If
    End If
End Sub


Sub EditorTab(control As IRibbonControl, ByRef visible)

visible = False
    thisuser = LCase(Environ("username"))
    
    If InStr(vbTextCompare, thisuser, "cressd02") > 0 Then visible = True
    If InStr(vbTextCompare, thisuser, "whitee06") > 0 Then visible = True
    

End Sub


Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    On Error Resume Next
    If myRibbon Is Nothing Then
        If RibbonID$ = "" Then RibbonID$ = GetRibbonID
        Set myRibbon = GetRibbon(RibbonID$)

        myRibbon.Invalidate
        Debug.Print "The Ribbon handle was lost, Hopefully this is sorted now by the GetRibbon Function?. You can remove this msgbox, I only use it for testing"
    
    Else
        myRibbon.Invalidate
    End If
    
    
End Sub






Sub SaveRibbonID(inRibbon$)
    If Len(inRibbon$) < 0 Then Exit Sub
    
    SaveFile$ = Environ("APPDATA")
    If Right(SaveFile$, 1) <> "\" Then SaveFile$ = SaveFile$ & "\"
    SaveFile$ = SaveFile$ & "RibbonID"
    
    On Error Resume Next
    'On Error GoTo 0
    If FileExists(SaveFile$) = True Then Kill SaveFile$

    Set fsoObject = CreateObject("Scripting.FileSystemObject")
    Set my_file = fsoObject.opentextfile(SaveFile$, 2, True)
    my_file.WriteLine inRibbon$
    my_file.Close

End Sub


Function GetRibbonID() As String
    
    SaveFile$ = Environ("APPDATA")
    If Right(SaveFile$, 1) <> "\" Then SaveFile$ = SaveFile$ & "\"
    SaveFile$ = SaveFile$ & "RibbonID"
    
    On Error Resume Next
    'On Error GoTo 0
    If FileExists(SaveFile$) = False Then Exit Function
    
    
    Set fsoObject = CreateObject("Scripting.FileSystemObject")
    Set my_file = fsoObject.opentextfile(SaveFile$, 1)
    GetRibbonID = my_file.readline
    my_file.Close


End Function



