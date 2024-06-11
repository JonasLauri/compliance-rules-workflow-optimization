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

    'With condition that 1st paragraph word begins with symbol "("
    'Iterate through 1st SD list to renumber 2nd (SD list No) word in paragraph
    'And count only True cycles
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "." Then Exit For
        If wPara.Range.Words(1) = "(" Then
            wPara.Range.Words(2) = a
            a = a + 1
            iCount = iCount + 1
        End If
        'For dublicate SD, adds space after 1st entity number
        If wPara.Range.Words(6).Characters(1) = "1" _
            Or wPara.Range.Words(6).Characters(1) = "2" _
            Or wPara.Range.Words(6).Characters(1) = "3" _
            Or wPara.Range.Words(6).Characters(1) = "4" _
            Or wPara.Range.Words(6).Characters(1) = "5" _
            Or wPara.Range.Words(6).Characters(1) = "6" _
            Or wPara.Range.Words(6).Characters(1) = "7" _
            Or wPara.Range.Words(6).Characters(1) = "8" _
            Or wPara.Range.Words(6).Characters(1) = "9" _
            Then wPara.Range.Words(4).InsertAfter " "
    Next wPara
    
    'Redeclare item range for arrays
    ReDim sdArray(iCount)
    ReDim numArray(iCount)
    iElement = 0
    
    'With same condition as previous
    'Make another 1st SD list iteration
    'Store 1st entity number and No of SD into seperate arrays
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "." Then Exit For
        If wPara.Range.Words(1) = "(" Then
            sdArray(iElement) = wPara.Range.Words(4)
            numArray(iElement) = wPara.Range.Words(2)
            iElement = iElement + 1
        End If
    Next wPara
    
    'Double interation through document objects
    '1. iteration is True when 1st paragraph word - symbol "."
    '1.1 iteration is True when 1st entity number of both SD lists are equal
    'Under True conditions 2nd list SD No is overwritten with 1st list SD No
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "." Then
            For iElement = 0 To iCount - 1
            'UPDATED LINE 2024-01-03: modified condition scope below to correctly numbering duplicate entity numbers
                If wPara.Range.Words(4) = sdArray(iElement) And numArray(iElement) <> "" Then
                    wPara.Range.Words(2) = numArray(iElement)
                    numArray(iElement) = ""
                    Exit For
                End If
            Next iElement
        End If
    Next wPara
    
    'For dublicate SD, code block of loop added to delete spaces after 1st entity number and comma
    For Each wPara In ActiveDocument.Paragraphs
        If wPara.Range.Words(1) = "." Then Exit For
        If wPara.Range.Words(1) = "(" And wPara.Range.Words(6).Characters(1) = "1" _
            Or wPara.Range.Words(6).Characters(1) = "2" _
            Or wPara.Range.Words(6).Characters(1) = "3" _
            Or wPara.Range.Words(6).Characters(1) = "4" _
            Or wPara.Range.Words(6).Characters(1) = "5" _
            Or wPara.Range.Words(6).Characters(1) = "6" _
            Or wPara.Range.Words(6).Characters(1) = "7" _
            Or wPara.Range.Words(6).Characters(1) = "8" _
            Or wPara.Range.Words(6).Characters(1) = "9" _
            Then
            wPara.Range.Words(4).Select
            With Selection
                .MoveRight
                .TypeBackspace
            End With
        End If
    Next wPara
                         
    'UPDATED 2022-11-29: added stability to code, numeration bug found, exceptions for dublicates noted;
    'UPDATED 2022-12-01: numeration problem for dublicate entities totaly solved
    'UPDATED 2024-01-03: corrected numeration of duplicate entity numbers
End Sub
