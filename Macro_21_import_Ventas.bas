Attribute VB_Name = "Macro_21_import_Ventas"
Sub Macro_21_import_Ventas()

'*FUNCCIONAL PARMAMETERS

'Ruta del fichero:D:\02 Work\201 - METRICAS de gestión

'Nombre del fichero:'1  – Ventas

'Nombre de la hoja:Ventas STD

'*GLOBAL VARIABLES AND CONTANTS
Dim wkbk1 As String ' WORKBOOK 1  – Ventas
Dim wkbk2 As String 'WORKBOOK 2 ALSO THISWORKBOOK - OUR DESTINATION & WHERE IS OUR MACRO
Dim rw1 As String
Dim rw2 As String
Dim rw2E As String
Dim sh1 As String 'SHEET VENTAS (ORIGINE)
Dim sh2 As String 'SHEET (DESTINATION)


'CHECK WE ARE IN THE WORKBOOK

wkbk2 = ThisWorkbook.Name

Windows(wkbk2).Activate
sh2 = ActiveSheet.Name
Sheets(sh2).Select


sh1 = "Ventas STD"
wkbk1 = "01 – VENTAS.xlsm"

' GOTO "VENTAS" WORKBOOK AND SELECT USED RANGE
 Windows(wkbk1).Activate
 Sheets(sh1).Select
 
 Range("T23").Select
 
 Do
 ActiveCell.Offset(1, 0).Select
 Loop Until IsEmpty(ActiveCell.Value)
 
 'MEMORIZE LAST ADDRESS IN COLUMN 'T'
 rw1 = ActiveCell.Row
 
 'SELECT REQUIRED RANGE
 Range(Cells(21, 19), Cells(rw1 - 1, 21)).Select
 Selection.Copy
 
 'GOTO TO DESTINATION
 Windows(wkbk2).Activate
 Sheets(sh2).Select
 
 'SELECT LAST AVAILABLE RANGE IN "A"
 Range("A17").Select
 Do
 ActiveCell.Offset(1, 0).Select
 Loop Until IsEmpty(ActiveCell.Value) And IsEmpty(ActiveCell.Offset(0, 1).Value)
 
 'MEMORIZE LAST ADDRESS IN COLUMN 'A'
 rw2 = ActiveCell.Row
 
 'PASTE TO DESTINATION
ActiveSheet.Paste
 
 'FORMAT OF THE SELECTED CELL
  With Selection.Font
        .Name = "Consolas"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
 End With
 
 
 'SET a 'X' on TOP LEFT CORNER
Cells(rw2, 1).Value = "x"
Range(Cells(rw2, 1), Cells(rw2, 1)).Select
 
 'FORMAT OF THE SELECTED CELL WITH THE 'x'
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
 End With
 
 
'FORMAT TOP ROW ABOVE TABLE - REMOVE BORDERS
Range(Cells(rw2, 1), Cells(rw2, 8)).Select
Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'COPY HEADERS
 Range("A2:E2").Select
 Selection.Copy
 Cells(rw2 + 1, 1).Select
 ActiveSheet.Paste
 
'DROP THE FOLLOWING ROW AFTER rw2
 Rows(rw2 + 2).Select
 Selection.Delete Shift:=xlUp
 
'COPY FAMILY U1 TO COLUMN -1
 Cells(rw2, 3).Select
 Selection.Copy
 Cells(rw2, 2).Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 Cells(rw2, 3).Select
 Selection.ClearContents
        
'PUT "x" TO LAST INOCUPIED ROW
 Cells(rw2, 1).Select
 Do
 ActiveCell.Offset(1, 0).Select
 Loop Until IsEmpty(ActiveCell.Value) And IsEmpty(ActiveCell.Offset(0, 1).Value)
 ActiveCell.Value = "x"
 
 'FORMAT OF THE SELECTED CELL
  With Selection.Font
        .Name = "Consolas"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
 End With
 With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
 End With
 
 'MEMORIZE LAST ADDRESS IN COLUMN 'A'
 rw2E = ActiveCell.Row
 
 'SELECT FORMULAS CALCUL
 Range("C3:E3").Select
 Selection.Copy
 'SELECT RANGE WHERE TO COPY FORMULAS
 Range(Cells(rw2 + 2, 3), Cells(rw2E - 1, 5)).Select
 'COPY FORMULAS
 Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

'SELECT FORMULAS FORMAT
 Range("C3:E3").Select
 Selection.Copy
 'SELECT RANGE WHERE TO COPY FORMULAS
 Range(Cells(rw2 + 2, 3), Cells(rw2E - 1, 5)).Select
 'COPY FORMULAS
 Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
 
        
    
 'SELECT FORMAT TAG
 Range("B1").Select
 Selection.Copy
    
 Cells(rw2, 2).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
 
End Sub

