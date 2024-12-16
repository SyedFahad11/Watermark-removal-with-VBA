Sub RemoveHyperlinkFromSlides()
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim urlToRemove As String
    Dim i As Long

    ' Define the URL to be removed
    urlToRemove = "https://www.knowledgegate.in/"  ' Replace with the exact URL you want to remove

    ' Loop through all slides in the active presentation
    For Each pptSlide In ActivePresentation.Slides
        ' Use a reverse loop to avoid collection modification issues
        For i = pptSlide.Shapes.Count To 1 Step -1
            Set pptShape = pptSlide.Shapes(i)

            ' Check for hyperlinks in text frames
            If pptShape.HasTextFrame Then
                If pptShape.TextFrame.HasText Then
                    If pptShape.ActionSettings(ppMouseClick).Hyperlink.Address = urlToRemove Then
                        ' Remove the hyperlink
                        pptShape.ActionSettings(ppMouseClick).Hyperlink.Delete
                    End If

                    ' Check if the text frame contains "Knowledge Gate Website"
                    If InStr(1, pptShape.TextFrame.TextRange.Text, "Knowledge Gate Website", vbTextCompare) > 0 Then
                        ' Delete the text frame
                        pptShape.Delete
                    End If
                End If
            End If

            
        Next i
    Next pptSlide

    ' Notify the user when the process is complete
    MsgBox "Specified hyperlinks and text frames removed from all slides!", vbInformation
End Sub

