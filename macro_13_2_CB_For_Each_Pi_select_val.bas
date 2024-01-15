Attribute VB_Name = "CB_For_Each_Pi_select_val"
Public pivot_sincro As Boolean
Public pivot_sincro_cell As Boolean


Sub main_macro_13()
If ActiveWorkbook.Worksheets("Ventas STD") Is ActiveSheet Then
Call CB_For_Each_Pi_select_val("pivot_table1", "pivot_table5", "Ventas STD", ThisWorkbook.Name, "Familia")
End If
If ActiveWorkbook.Worksheets("Ventas EOY") Is ActiveSheet Then
Call CB_For_Each_Pi_select_val("pivot_table1", "pivot_table5", "Ventas EOY", ThisWorkbook.Name, "Familia")
End If
Application.Calculation = xlCalculationAutomatic
End Sub
 

Sub CB_For_Each_Pi_select_val(master_pt As String, slave_pt As String, hoja As String, wkbk As String, fld As String)
ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

'PICK WORKBOOK
Windows(wkbk).Activate

'SELECT SHEETS PIVOTS
Sheets(hoja).Select

'PREVENT THE MACRO TO RUN TWICE
For iter = 1 To 2
If iter > 1 Then Exit Sub

'VARIABLES SETTING
Dim pt As PivotTable
Set pt = ActiveSheet.PivotTables(master_pt)
Dim pf As PivotField
Set pf = pt.PivotFields(fld)
Dim itm As PivotItem
Dim count_itm_true As Integer 'COUNT PIVOT ITEM IN PIVOT TABLE
Dim new_selection As String  ' VALUE IN PIVOT SELECTED
new_selection = Cells(2, 2).Value
Dim prev_selection As String      ' VALUE IN PIVOT SELECTED PREVIOUSLY
prev_selection = Cells(3, 2).Value
Dim bnew  As Boolean 'BOOLEAN TRUE IF NEW VALUE SELECTED
Dim bprev As Boolean 'BOOLEAN TRUE IF OLD VALUE SELECTED
bnew = False
bprev = False


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

'RESET PIVOT ITEM COUNTER
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
                For i = 1 To 2 'FIX BUG ON TRUE STATEMENT
                If i = 1 Then
                itm.Visible = True
                End If
            Next i
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

'PREVENT THE MACRO TO RUN TWICE
Next iter

pt.ManualUpdate = False

Application.Calculation = xlCalculationAutomatic
pf.PivotCache.Refresh
Calculate
End Sub
