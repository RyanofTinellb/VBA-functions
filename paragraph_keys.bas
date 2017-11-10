Sub MoveParagraphUp()
    MoveParagraph goUp:=True
End Sub

Sub MoveParagraphDown()
    MoveParagraph goUp:=False
End Sub

Sub CutParagraph()
    Selection.Paragraphs(1).Range.Cut
End Sub

Sub MoveParagraph(goUp As Boolean)
    CutParagraph
    If goUp Then
        Selection.MoveUp Unit:=wdParagraph, Count:=1
    Else
        Selection.MoveDown Unit:=wdParagraph, Count:=1
    End If
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveUp Unit:=wdParagraph, Count:=1
End Sub
