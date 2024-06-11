Attribute VB_Name = "m_AutoRulesDocumentation"
Sub Auto_Rules_Documentation()
'NOTES BEFORE USING THE MACRO:
'!!!To run properly the macro both relevant workbooks must be opened!!!
'!!!Run this macro only on Fridy (Rule implementation day)!!!
'There is an option to locate your new rules excel file and open it automatically

Dim wsCopy As Worksheet
Dim wsDest As Worksheet
Dim wsCopyRow As Range
Dim lCopyLastRow As Long
Dim lDestLastRow As Long
    
'Set variables for copy and destination sheets
Set wsCopy = Workbooks("Rules.xlsx").Worksheets("Sheet1")
Set wsDest = Workbooks("False Positive reduction rules.xlsm").Worksheets("Sheet1")

'Turn original workbook as active in win screen
wsDest.Activate

'0. Color 1st row in blue
wsCopy.Rows(2).Interior.Color = RGB(0, 176, 240)

'1. Color a newly added entity in dark red & Make data prefills in empty cells
'1.1 First enter data using input boxes and store in a variable
Name1 = InputBox("Enter credentials who performed 4 Eyes Control?", "4 Eyes Control Performed By", "BD8882/BE9260")
Name2 = InputBox("Enter credentials who approved the rules?", "Approved by", "QSA/EGKAC/VSL")
Name3 = InputBox("Enter credentials who implemented the rules?", "Implemented by", "SARMA/JONLAU")

For Each wsCopyRow In wsCopy.UsedRange.Rows
    If Not wsCopyRow.Find("This is a newly added entity") Is Nothing Then
        wsCopyRow.Font.Color = RGB(192, 0, 0)
    End If
    wsCopyRow.Range("I2:J2").Value = Date
    wsCopyRow.Range("K2").Value = Date + 3
    wsCopyRow.Range("H2").Value = Name1
    wsCopyRow.Range("L2").Value = Name2
    wsCopyRow.Range("M2").Value = Name3
Next wsCopyRow
'1.2 Color exceptions in green & add cell comments containing entity description
For Each wsCopyRow In wsCopy.UsedRange.Rows
    If Not wsCopyRow.Find("This is an exception.") Is Nothing Then
        With wsCopyRow.Range("B1")
            .Interior.Color = RGB(146, 208, 80)
            .AddComment (wsCopyRow.Range("N1").Value)
        End With
    End If
Next wsCopyRow

'2. Find last used row in the copy range based on data in column B
lCopyLastRow = wsCopy.Cells(wsCopy.Rows.Count, "B").End(xlUp).Row
  
'3. Find first blank row in the destination range based on data in column B
'Offset property moves down 1 row
lDestLastRow = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).Offset(1).Row

'4. Copy & Paste Data
wsCopy.Range("B2:N" & lCopyLastRow).Copy _
  wsDest.Range("B" & lDestLastRow)
  
'5. Closes new rules workbook
Workbooks("Rules.xlsx").Close SaveChanges:=False

'Update 2022-11-08: Additional funtionality was added, reference to 1.2
End Sub
