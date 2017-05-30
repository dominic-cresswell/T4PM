Attribute VB_Name = "W_DownloadCode"

Option Private Module


Sub PushReadDataToDocument(dummy$)

    For zzz = 0 To 9999
            FieldName1$ = ProjectReadDataArray(zzz, 0)
            
            If FieldName1$ = "" Then Exit For
            
            FieldName1$ = Replace(FieldName1$, "_n0", "_null")
            FieldName2$ = FieldName1$
            
            FieldName1$ = "T4PM_S_W_" & FieldName1$
            FieldName2$ = "T4PM_S_R_" & FieldName2$
            
            FieldData$ = ProjectReadDataArray(zzz, 1)
            FieldStamp$ = ProjectReadDataArray(zzz, 2)
            

                On Error Resume Next

                sh.Range(FieldName1$) = FieldData$
                sh.Range(FieldName2$) = FieldData$

                FindText$ = FieldName1$
                ReplacedText$ = FieldData$
                GoSub DoShapeReplace:
                
                FindText$ = FieldName2$
                ReplacedText$ = FieldData$
                GoSub DoShapeReplace:
                
                
               For locale = 1 To 11
           
                FindText$ = FieldName1$
                ReplacedText$ = FieldData$
                GoSub DoReplace:
                
                FindText$ = FieldName2$
                ReplacedText$ = FieldData$
                GoSub DoReplace:
               Next
           
    Next




MopUp:
' mop up the leftovers

    With ActiveDocument.StoryRanges.Item(wdMainTextStory).Find
        .Text = "<T4PM*_*_*_*_*>"
        .MatchWildcards = True
        .Replacement.Text = "T4PM"
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.StoryRanges.Item(wdMainTextStory).Find
        .Text = "<<T4PM>>"
        .MatchWildcards = False
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
Exit Sub


DoShapeReplace:

 ' do shapes
   For Each sh In ActiveDocument.Shapes
        GetText$ = sh.TextFrame.TextRange
        sh.TextFrame.TextRange.Text = Replace(GetText$, "<<" & FindText$ & ">>", ReplacedText$)
    Next
   
    Return
    
DoReplace:
    ' do everything else
    With ActiveDocument.StoryRanges.Item(locale).Find
             .Text = "<<" & FindText$ & ">>"
             .Replacement.Text = ReplacedText$
             .Forward = True
             .Wrap = wdFindContinue
             .Format = False
             .MatchCase = False
             .MatchWholeWord = False
             .MatchWildcards = False
             .MatchSoundsLike = False
             .MatchAllWordForms = False
             .Execute Replace:=wdReplaceAll
     End With
     


     
Return
End Sub
