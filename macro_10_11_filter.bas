Attribute VB_Name = "ALAIN"


'MACRO ALIGNED WITH FILTER CHANGE
Sub copyLastfilVAL3_macro_10()

Dim rngb As Range
Dim cels As Range

'SPOT LAST CELL COLUMN 'B'
Cells(5, 3).Select
Do Until IsEmpty(ActiveCell.Offset(1, 0)) Or ActiveCell.Value = "Z-Test"
ActiveCell.Offset(1, 0).Select
Loop

Set rnga = Range(Cells(5, 3), Cells(ActiveCell.Row, 3))

Set rngb = ActiveSheet.Range(rnga.Address).SpecialCells(xlCellTypeVisible) 'METHOD THAT SELECT THE VISIBLE RANGE, '!!!!' THE RANGE MUST BE CORRECTLY SELECTED WITH NO BLANKS

'"!" WE LOOP THROUGH THE VALUES OF THE RANGE 'rngb' IN ORDER TO CHOSE LAST THE VALUE VISIBLE END FROM THE LIST "!!!" ENABLE TO NOT CRASH THE WORKBOOK

For Each cels In rngb
Cells(1, 4) = cels 'PRINT RESULT IN 'D1'

If cels = "z-Test" Then Cells(1, 4) = Null
'ActiveSheet.ListObjects("Table5").Range.AutoFilter Field:=2

Next cels


End Sub

'MACRO ALIGNED WITH FILTER CHANGE
Sub PickUpVal_macro_11()

Dim rngb As Range
Dim cels As Range
Dim chk As Boolean

'SPOT LAST CELL COLUMN 'B'
Cells(5, 3).Select
Do Until IsEmpty(ActiveCell.Offset(1, 0)) Or ActiveCell.Value = "Z-Test"
ActiveCell.Offset(1, 0).Select
Loop

'START AT RANGE C1
Set rnga = Range(Cells(5, 3), Cells(ActiveCell.Row, 3))
'STORE RANGE C1
Set rngb = ActiveSheet.Range(rnga.Address)

'"!" WE LOOP THROUGH THE VALUES OF THE RANGE 'rngb'
For Each cels In rngb

'IF VALUE IN C1 IS THE COLUMN C5:C40
If cels = Cells(1, 3).Value Then
    chk = True
   'COPY THE VALUE FROM M1:U1
   Range("M1:U1").Copy
   Cells(cels.Row, 13).Select
   'PASTE SELECTION
   ActiveSheet.Paste
   
   'GO TO ROW WHERE THE C1 VALUE IS
   Cells(cels.Row, 13).Select
   
   'CLEAN COPIED RANGE
   Do Until IsEmpty(ActiveCell.Offset(0, 1))
   ActiveCell.Offset(0, 1).Select
        Select Case True
        ' SET NULL IF NOT A NUMBER
        Case IsNumeric(ActiveCell.Value) = False
        ActiveCell.Value = Null
        ' SET NULL IF EQUAL ZERO
        Case ActiveCell.Value = 0
        ActiveCell.Value = Null
        Case Else
        End Select
   Loop
Else
End If

Next cels

If chk = False Then MsgBox "Value in C1 is Not in the C Column!"

End Sub



'NOT FISNISHED  ... BUT MAY HELP LATER
Function FXfirstfilVAL5(rnga As Range) As String


Dim cels As Range
'Set rngb = ActiveSheet.Range(Cells(5, 3), Cells(90, 3)).SpecialCells(xlCellTypeVisible) ' ADD THE RANGE B5:B90 TO THE OBJET 'rngb'
Set rngb = ActiveSheet.Range(rnga.Address).SpecialCells(xlCellTypeVisible) ' ADD THE RANGE B5:B90 TO THE OBJET 'rngb'

For i = 1 To rngb.Count 'HOW MANY VALUE IN THE RANGE B5:B90

    'LOOP ALL THE RECORDS IN OBJECT
    For Each cels In rngb
        If cels = Cells(1, 4).Value Then GoTo exitF Else 'IF EQUAL 'D1' AS CONDITION BECAUSE xlCellTypeVisible METHOD DOES NOT WORK WELL ON A FUNCTION
        
    Next cels
Next i

exitF:
FXfirstfilVAL5 = cels.Value  '& rnga.Cells(i, 3) 'PRINT RESULT IN 'E1'
'FXfirstfilVAL5 = rnga.Address
End Function
