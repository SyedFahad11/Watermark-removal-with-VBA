Sub RemoveTextFromAllPages()
    Dim rng As Range
    
    ' Turn off screen updating for faster execution
    Application.ScreenUpdating = False
    
    ' Loop through all pages in the document
    For Each rng In ActiveDocument.StoryRanges
        Do
            ' Find and remove the specified text
            With rng.Find
                .Text = "Gate.appliedcourse.com"
                .Replacement.Text = ""
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
            
            With rng.Find
                .Text = "Ph: +91 844-844-0102"
                .Replacement.Text = ""
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
            
            ' Move to the next linked story range (e.g., headers, footers)
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Text removal complete!", vbInformation
End Sub
