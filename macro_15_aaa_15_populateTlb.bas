Attribute VB_Name = "aaa_15_populateTlb"
Sub macro_15_populateTlb()
Attribute macro_15_populateTlb.VB_ProcData.VB_Invoke_Func = " \n14"
Application.ScreenUpdating = False
'PUT THE SHEET NAME DESTINATION
Const wkb2 As String = "macro_15_16_20231224.xlsm" '<== Pon tu worbook de destinacíon aqui
Const wks2 As String = "Oo Stock" '<== = Pon tu hoja de destinacíon aqui

'RECORD THE CURRENT WORKBOOK + WORKSHEET LOCATION
Dim wkb1 As Workbook
Dim wks1 As Worksheet
Dim addrr_1 As String

'COPY THE SELECTED CELL
 Set wkb1 = ActiveWorkbook
 Set wks1 = ActiveSheet
 addrr_1 = ActiveCell.Address

 'EXIT MACRO IF POSITONED ON DESTINATION SHEET
 If wks1.Name = wks2 Then Exit Sub
 
'GOTO TO TBL SHEET
 Workbooks(wkb2).Activate
    Sheets(wks2).Activate
    
'INSERT BLANK ROW
    Rows("2:2").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'RETURN TO SELECTION
    wkb1.Activate
    wks1.Activate
    Range(addrr_1).Select
    Selection.Copy
    
'GOTO TO TBL SHEET
Workbooks(wkb2).Activate
    Sheets(wks2).Activate
    
'COPY SELECTION TO UNOCUPIED ROW
    Cells(2, 1).Select
    ActiveSheet.Paste
    
 'APPEND FIELDS
    ActiveCell.Offset(0, 1) = Date
    ActiveCell.Offset(0, 2).FormulaR1C1 = "=RANDBETWEEN(0,6)^2" ' <== Reemplaca esta formula de test con tu formula VLOOKUP
                                                                    'Es decir que deberia ser esta formula :  "=VLOOKUP(RC,'D:\02 Work\201- METRICAS de gestión\[100 - LUT Familias - Stock Vtas Devols.xlsx]LUT familia'!R1C1:R1003C3,2,FALSE)"
'FORMAT BORDERS
Range("A1:C2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
'RETURN TO SELECTION
    wkb1.Activate
    wks1.Activate
    Range(addrr_1).Select
Application.ScreenUpdating = True
End Sub
