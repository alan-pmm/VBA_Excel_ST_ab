Attribute VB_Name = "Macro_22_end_sum"
Sub Macro_22_end_sum()
Dim wkbk2 As String 'WORKBOOK 2 ALSO THISWORKBOOK - OUR DESTINATION & WHERE IS OUR MACRO
Dim sh2 As String 'SHEET (DESTINATION)
Dim rw2 As String

'CHECK WE ARE IN THE WORKBOOK

wkbk2 = ThisWorkbook.Name

Windows(wkbk2).Activate
sh2 = ActiveSheet.Name
Sheets(sh2).Select


'SELECT ROW WHERE TO START
Range("A17").Select
'Cells(17, 1).Select


Do
ActiveCell.Offset(1, 0).Select
Loop Until IsEmpty(ActiveCell.Value) And IsEmpty(ActiveCell.Offset(0, 1).Value)

'MEMORIZE LAST ADDRESS IN COLUMN 'A'
rw2 = ActiveCell.Row

Range("A5:E11").Select
Selection.Copy

Cells(rw2, 1).Select
ActiveSheet.Paste

Cells(rw2, 3).Select
Selection.Clear

' -- SCENARIO 1 STATIC RANGE FOR SUM
ActiveCell.FormulaR1C1 = "=SUM(R[-10]C:R[-4]C)"

' -- SCENARIO 2 DECOMMISONATED
'CALL A VBA CLASS FOR DYNAMIC SUM RANGE - WILL SUM ALL THE C COLUMN
'ActiveCell.Value = Application.WorksheetFunction.Sum(Range(Cells(18, 3), Cells(rw2 - 2, 3)))

'RETURN TO THE TOP OF THE SHEET AND DELETE ORIGINAL HEADERS
Rows("1:13").Select
Selection.Delete Shift:=xlUp

End Sub

