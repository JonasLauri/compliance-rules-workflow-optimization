Attribute VB_Name = "TypePara_MondayLetter"
Sub TypePara_MondayLetter()
    Dim wPara As Paragraph
    
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "Once" Then Exit For
        If wPara.Range.Words(1) = "." Then
            wPara.Range.Words(8).Select
            With Selection
                .MoveRight
                .TypeParagraph
            End With
        End If
    
    Next wPara
    
    ActiveDocument.Paragraphs(1).Range.Words(1).Select
    
End Sub
