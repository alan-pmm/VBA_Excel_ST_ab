Public new_selection As String  'VALUE IN PIVOT SELECTED
Public prev_selection As String 'VALUE IN PIVOT SELECTED PREVIOUSLY
Public count_itm_true As Integer 'COUNT PIVOT ITEM IN PIVOT TABLE

Sub main_macro_13_STD()


If off = True Then Exit Sub

'PIVOT 2 STD
' CHECK IF RANGE "P5" = "ON" TO ACTIVATE OR DESACTIVATE THE MACRO
If ActiveWorkbook.Worksheets("Ventas STD").Range("P5").Value = "ON" Then
Else
Exit Sub
End If

' CHECK IF ACTIVESHEET IS 'VENTAS STD'
If ActiveWorkbook.Worksheets("Ventas STD") Is ActiveSheet Then
Call CA_ctrl_pt_macro_13("pivot_table1", "Ventas STD", ThisWorkbook.Name, "Familia")
Else
End If

'CHECK IF PIVOT 1 HAS ONE VALUE SELECTED
If count_itm_true > 1 Then
Exit Sub
Else
Call CB_syncr_pt_macro_13("pivot_table5", "Ventas STD", ThisWorkbook.Name, "Familia")
End If

'PIVOT 3 STD
' CHECK IF RANGE "P6" = "STD_pt3" TO ACTIVATE OR DESACTIVATE THE MACRO
If ActiveWorkbook.Worksheets("Ventas STD").Range("P6").Value = "STD_pt3" Then
Else
Exit Sub
End If

'CHECK IF PIVOT 1 HAS ONE VALUE SELECTED
If count_itm_true > 1 Then
Exit Sub
Else
Call CB_syncr_pt_macro_13("pivot_table5", "Pivot Stock", "STOCK.xlsm", "Familia")
End If

'COME BACK TO ''VENTAS''
ThisWorkbook.Activate
Sheets("Ventas STD").Select
End Sub

Sub main_macro_13_EOY()

'PIVOT 2 EOY
' CHECK IF RANGE "P5" = "ON" TO ACTIVATE OR DESACTIVATE THE MACRO
If ActiveWorkbook.Worksheets("Ventas STD").Range("P5").Value = "ON" Then
Else
Exit Sub
End If

' CHECK IF ACTIVESHEET IS 'VENTAS EOY'
If ActiveWorkbook.Worksheets("Ventas EOY") Is ActiveSheet Then
Call CA_ctrl_pt_macro_13("pivot_table1", "Ventas EOY", ThisWorkbook.Name, "Familia")
Else
End If

'CHECK IF PIVOT MASTER HAS ONE VALUE SELECTED
If count_itm_true > 1 Then
Exit Sub
Else
Call CB_syncr_pt_macro_13("pivot_table5", "Ventas EOY", ThisWorkbook.Name, "Familia")
End If

'PIVOT 3 EOY
' CHECK IF RANGE "P6" = "EOY_pt3" TO ACTIVATE OR DESACTIVATE THE MACRO
If ActiveWorkbook.Worksheets("Ventas STD").Range("P6").Value = "EOY_pt3" Then
Else
Exit Sub
End If

'CHECK IF PIVOT 1 HAS ONE VALUE SELECTED
If count_itm_true > 1 Then
Exit Sub
Else
Call CB_syncr_pt_macro_13("pivot_table5", "Pivot Stock", "STOCK.xlsm", "Familia")
End If

'COME BACK TO ''VENTAS''
ThisWorkbook.Activate
Sheets("Ventas EOY").Select
End Sub
 

Sub CA_ctrl_pt_macro_13(master_pt As String, hoja As String, wkbk As String, fld As String)

' INPUT VALUES FOR PIVOTITEM
new_selection = Cells(2, 2).Value
prev_selection = Cells(3, 2).Value
Application.Calculation = xlCalculationManual


'PICK WORKBOOK
Windows(wkbk).Activate

'SELECT SHEETS PIVOTS
Sheets(hoja).Select

'VARIABLES SETTING
Dim pt As PivotTable
Set pt = ActiveSheet.PivotTables(master_pt)
Dim pf As PivotField
Set pf = pt.PivotFields(fld)
Dim itm As PivotItem

'RESET PIVOT ITEM COUNTER
count_itm_true = 0

'ENSURE ONE ITEM IS SELECTED IN PIVOT MASTER
pt.ManualUpdate = True
For Each itm In pf.PivotItems
If itm.Visible = True Then
count_itm_true = count_itm_true + 1
End If
If count_itm_true > 1 Then
MsgBox "Select only one value in Master pivot. The macro will stop here"
Application.Calculation = xlCalculationAutomatic
Exit Sub
End If
Next itm
'CODE EQUIVALENT
'If InStr(new_selection, "Multipl") > 0 Then
 '       MsgBox "Select only one value. The macro will stop here"
 '       Exit Sub
'End If
Application.Calculation = xlCalculationAutomatic
End Sub

Sub CB_syncr_pt_macro_13(slave_pt As String, hoja As String, wkbk As String, fld As String)

Dim pt As PivotTable
Set pt = ActiveSheet.PivotTables(slave_pt)
Dim pf As PivotField
Set pf = pt.PivotFields(fld)
Dim itm As PivotItem

Dim bnew  As Boolean 'BOOLEAN TRUE IF NEW VALUE SELECTED
Dim bprev As Boolean 'BOOLEAN TRUE IF OLD VALUE SELECTED
bnew = False
bprev = False
Dim count_itm_true As Integer 'COUNT PIVOT ITEM IN PIVOT TABLE

'PICK WORKBOOK
Windows(wkbk).Activate

'SELECT SHEETS PIVOTS
Sheets(hoja).Select

count_itm_true = 0
'SET PIVOT SLAVE
Set pt = ActiveSheet.PivotTables(slave_pt)
Set pf = pt.PivotFields(fld)
'SWITH 1 - SET CONDITION FOR A VALID SELECTION
Application.Calculation = xlCalculationManual
For Each itm In pf.PivotItems
If itm.Visible = True Then
    Select Case itm.Name
        'CONDITION1
        Case new_selection
        bnew = True
        count_itm_true = count_itm_true + 1
        'CONDITION2
        Case prev_selection
        bprev = True
        count_itm_true = count_itm_true + 1
        Case Else
        count_itm_true = count_itm_true + 1
    End Select
Else
End If
Next itm
Application.Calculation = xlCalculationAutomatic
'CODE EQUIVALENT = 'pct = ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Familia1").VisibleSlicerItems.Count

'SWITH 2 - ACCORDING CONDITION, ATTRIBUTE SCENARIO 1, 2 AND/OR 3
'CONDITION1
If (bprev = True And count_itm_true = 1) Then
GoTo 2
Else
End If
'CONDITION2
If (bnew = True And count_itm_true = 1) Then
GoTo 3
End If

'RESET THE FILTERS OF THE PIVOT SLAVE
1:
pt.ManualUpdate = True

'RESET FILTERS TO ALL
pf.ClearAllFilters
   
'ATTACK SELECTION ON PIVOT TABLE X
Application.Calculation = xlCalculationManual
For Each itm In pf.PivotItems
        Select Case itm.Name
            Case new_selection ' la posicion de la celda donde se copia el valor de la familia inicial
                For I = 1 To 2 'FIX BUG ON TRUE STATEMENT
                If I = 1 Then
                itm.Visible = True
                End If
            Next I
            Case Else
                On Error Resume Next
                itm.Visible = False
                On Error GoTo 0
        End Select
    Next itm
Application.Calculation = xlCalculationAutomatic
'SWITH PIVOTITEM BETWEEN OLD SELECTION TO NEW SELECTION
2:
On Error GoTo 1
ActiveSheet.PivotTables(slave_pt).PivotFields(fld).PivotItems(new_selection).Visible = True

'ATTACK SELECTION ON PIVOT TABLE SLAVE
If prev_selection <> new_selection Then
ActiveSheet.PivotTables(slave_pt).PivotFields(fld).PivotItems(prev_selection).Visible = False
Else
End If

'COPY THE NEW SELECTION VALUE TO THE OLD SELECTION ALLOCATION
3:
Cells(2, 2).Copy
Cells(3, 2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


pt.ManualUpdate = False

Application.Calculation = xlCalculationAutomatic
ActiveSheet.PivotTables(slave_pt).PivotCache.Refresh


Calculate
End Sub