Attribute VB_Name = "AA_Macro_20"
Sub Macro_20_rule_out_0_NULL_values()
'
' Macro_20_rule_out_0_NULL_values
'
Dim ca As String 'CELL ADDRESS IN SHEET "PIVOT"
Dim cb As String 'CELL ADDRESS IN SHEET "RAW DATA"
Dim sh1 As String
Dim sh2 As String
sh1 = "Pivot" '   < -- nombre de tu hoja pivot
sh2 = "RawData" ' < -- nombre de tu hoja raw data

'SCAN I
Sheets(sh1).Select
Range("I17").Select

Do
ActiveCell.Offset(1, 0).Select
ca = ActiveCell.Address

Select Case True
Case ActiveCell.Value > 0
    'COPY RANK FROM SHEET "PIVOT" TO SHEET "RAW DATA"
        Selection.Copy
        Sheets(sh2).Select
    'SELECT THE TOP LOCATION FROM SHEET "RAW DATA"
        Range("F1").Select
    'LOOP UNTIL UNOCCUPIED CELL FROM SHEET "RAW DATA" AND COPY THE "I" COLUMN TAKEN FROM SHEET "PIVOT" TO SHEET "RAW DATA" -> COLUMN "F"
        Do
        ActiveCell.Offset(1, 0).Select
        Loop Until IsEmpty(ActiveCell.Value)
        ActiveSheet.Paste
        cb = ActiveCell.Address
    'RETURN TO SHEET "PIVOT" AND PICK UP THE VALUE IN COLUMN "B"
        Sheets(sh1).Select
        Range(ca).Select
        ActiveCell.Offset(0, -7).Select
        Selection.Copy
     'GOTO SHEET "RAW DATA" AND COPY THE COLUMN "B" FROM SHEET "PIVOT" TO THE SHEET "RAW DATA" -> COLUMN "A"
        Sheets(sh2).Select
        Range(cb).Select
        ActiveCell.Offset(0, -5).Select
        ActiveSheet.Paste
      'RETURN TO SHEET "PIVOT"
        Sheets(sh1).Select
        Range(ca).Select
        
      'RETURN TO SHEET "PIVOT" AND DO NOTHING
Case Len(ActiveCell.Value) = 0
        Sheets(sh1).Select
        Range(ca).Select
        
       'RETURN TO SHEET "PIVOT" AND DO NOTHING
Case Else
        Sheets(sh1).Select
        Range(ca).Select

End Select

Loop Until IsEmpty(ActiveCell.Offset(0, -7).Value)

End Sub

