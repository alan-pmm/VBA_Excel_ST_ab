Attribute VB_Name = "AB_UnColorRow_address_1"
Sub AB_UnColorRow_address_1()

'PICK UP VARIABLE ADRESS
Dim ur As Integer
Dim uc As Integer
Cells(1, 80).Select
ur = Cells(1, 80).Value  'CB1
Cells(1, 81).Select
uc = Cells(1, 81).Value  'CC1

Range(Cells(ur, 1), Cells(ur, uc)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range(Cells(ur + 1, 1), Cells(ur, uc)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range(Cells(ur + 2, 1), Cells(ur, uc)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range(Cells(ur + 3, 1), Cells(ur, uc)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Range(Cells(ur + 4, 1), Cells(ur, uc)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub
