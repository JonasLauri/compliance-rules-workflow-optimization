Attribute VB_Name = "AutoSD_ListsNumeration"
Sub AutoSD_ListsNumeration()
    'Numeration arrangement of two SD lists after rule ideation process is complete
    
    Dim wPara As Paragraph
    Dim sdArray() As String
    Dim numArray() As String
    Dim a As Integer
    Dim iCount As Integer
    Dim iElement As Integer
    
    a = 1
    iCount = 0

    'With condition that 1st paragraph word begins with symbol "(" _
    'Iterate through 1st SD list to renumber 2nd paragraph words _
    'And count only True cycles
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "(" Then
            wPara.Range.Words(2) = a
            a = a + 1
            iCount = iCount + 1
        End If
    Next wPara
    
    'Redeclare item range for arrays
    ReDim sdArray(iCount)
    ReDim numArray(iCount)
    iElement = 0
    
    'With same condition as previous _
    'Make another 1st SD list iteration _
    'Store 1st entity number and No of SD into seperate arrays
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "(" Then
            sdArray(iElement) = wPara.Range.Words(4)
            numArray(iElement) = wPara.Range.Words(2)
            iElement = iElement + 1
        End If
    Next wPara
    
    'Double interation through document objects _
    '1. iteration is True when 1st paragraph word - symbol "." _
    '1.1 iteration is True when 1st entity number of both SD lists are equal _
    'Under True conditions 2nd list SD No is overwritten with 1st list SD No
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "." Then
            For iElement = 0 To iCount
                If wPara.Range.Words(4) = sdArray(iElement) Then
                    wPara.Range.Words(2) = numArray(iElement)
                End If
            Next iElement
        End If
    Next wPara
                           
End Sub
