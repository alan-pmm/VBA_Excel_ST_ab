Attribute VB_Name = "AA_ColorRow_address_1"
Sub AA_ColorRow_address_1()
Attribute AA_ColorRow_address_1.VB_ProcData.VB_Invoke_Func = " \n14"

Dim r As Integer
r = ActiveCell.Row
'Call AB_UnColorRow_address_1.AB_UnColorRow_address_1
Const c As Integer = 31 ' COLUMN NUMBER WHETE TO STOP SELECTION EX 16 = "P" AND 29 = "AC"

'STORE VARIABLE ADRESS
Cells(1, 80).Value = r
Cells(1, 81).Value = c

Range(Cells(r, 1), Cells(r, c)).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 13421823
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


Range(Cells(r + 1, 1), Cells(r + 1, c)).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


Range(Cells(r + 2, 1), Cells(r + 2, c)).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 10284031
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

Range(Cells(r + 3, 1), Cells(r + 3, c)).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15189683
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
    
Range(Cells(r + 4, 1), Cells(r + 4, c)).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 13224393
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


    
End Sub

