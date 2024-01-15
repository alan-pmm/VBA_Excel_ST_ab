Attribute VB_Name = "CA_For_Each_Pi_select_val"


Sub main()
'Call CA_For_Each_Pi_select_val("pivot_table5", "Ventas STD", ThisWorkbook.Name, "Familia")
Call CB_For_Each_Pi_select_val("pivot_table5", "Ventas STD", ThisWorkbook.Name, "Familia")
'Call CA_For_Each_Pi_select_val("PivotTable3", "p1", ThisWorkbook.Name)
'Call CA_For_Each_Pi_select_val("PivotTable7", "pivotest3", "us-500-20221015-third-pivot.xlsm")
End Sub
 
 


Sub CA_For_Each_Pi_select_val(x As String, hoja As String, wkbk As String, fld As String)
ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationAutomatic
Windows(wkbk).Activate
Dim pt As PivotTable
Set pt = ActiveSheet.PivotTables(x)
Dim pf As PivotField
Set pf = pt.PivotFields(fld)
Dim itm As PivotItem
Dim i As Long
Dim zod As String
o = Cells(2, 2).Value
 
'SELECT SHEETS PIVOTS
Sheets(hoja).Select
 
' SPEED UP SELECTION
pt.ManualUpdate = True
 
 
'RESET FILTERS TO ALL
'On Error Resume Next
pf.ClearAllFilters
'On Error GoTo 0



'ATTACK SELECTION ON PIVOT TABLE X
Application.Calculation = xlCalculationManual
For Each itm In pf.PivotItems
 
    Select Case itm.Name
            
    Case o ' la posición de la celda donde se copia el valor de la familia inicial
    For i = 1 To 2 'FIX BUG ON TRUE STATEMENT
       If i = 1 Then itm.Visible = True Else
       Next i
        
    Case Else
    'On Error Resume Next
    itm.Visible = False
    'On Error GoTo 0
        
    End Select
        
Next itm
pt.ManualUpdate = False

Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
'pt.PivotCache.Refresh
End Sub


Sub CB_For_Each_Pi_select_val(x As String, hoja As String, wkbk As String, fld As String)
ScreenUpdating = False

Windows(wkbk).Activate

Dim pt As PivotTable
Set pt = ActiveSheet.PivotTables(x)
Dim pf As PivotField
Set pf = pt.PivotFields(fld)
Dim itm As PivotItem
Dim i As Long
Dim newsel As String
newsel = Cells(2, 2).Value
Dim prevsel As String
prevsel = Cells(3, 2).Value
Dim pct As Long
Dim bnew  As Boolean
Dim bprev As Boolean
bnew = False
bprev = False
'SELECT SHEETS PIVOTS
Sheets(hoja).Select

If newsel = "(Multiple Items)" Then
MsgBox "Select only one value. The macro will stop here"
Exit Sub
End If

'SPEED UP SELECTION
'pt.RefreshTable

'ActiveSheet.PivotTables("pivot_table5").PivotFields("Familia").VisibleItems.Count
'ActiveWorkbook.SlicerCaches(1).VisibleSlicerItems.Count



pct = ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_Familia1").VisibleSlicerItems.Count
If pct > 2 Then GoTo 1

'SWITH 1
'For pct = 1 To 2
For Each itm In pf.PivotItems
If itm.Visible = True Then
    
    Select Case itm.Name
        'CONDITION1
        Case prevsel
        bprev = True
        'CONDITION2
        Case newsel
        bnew = True
    End Select
End If
Next itm
'Next pct


'SWITH 2
'CONDITION1
If (bprev = True And pct = 1) Then
GoTo 2
Else
End If
'CONDITION2
If (bnew = True And pct = 1) Then
Exit Sub
Else
End If


    
    
1:
pt.ManualUpdate = True


    'RESET FILTERS TO ALL
    'On Error Resume Next
    pf.ClearAllFilters
    'On Error GoTo 0
    
    'ATTACK SELECTION ON PIVOT TABLE X
    Application.Calculation = xlCalculationManual
    
    For Each itm In pf.PivotItems
     
        Select Case itm.Name
                
        Case newsel ' la posición de la celda donde se copia el valor de la familia inicial
        For i = 1 To 2 'FIX BUG ON TRUE STATEMENT
           If i = 1 Then itm.Visible = True Else
        Next i
            
        Case Else
        On Error Resume Next
        itm.Visible = False
        On Error GoTo 0
            
        End Select
            
    Next itm


2:


'RESET FILTERS TO ALL
On Error GoTo 1
ActiveSheet.PivotTables(x).PivotFields(fld).PivotItems(newsel).Visible = True

'ATTACK SELECTION ON PIVOT TABLE X

3:

If prevsel <> newsel Then
ActiveSheet.PivotTables(x).PivotFields(fld).PivotItems(prevsel).Visible = False
Else
End If

Cells(2, 2).Copy
Cells(3, 2).Select
ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

pt.ManualUpdate = False
Application.Calculation = xlCalculationAutomatic

End Sub

Sub CountVisible()
Dim ws As Worksheet, pt As PivotTable, pf As PivotField
Set ws = ActiveSheet
Set pt = ws.PivotTables("pivot_table5")
Set pf = pt.PivotFields("Familia")
MsgBox "Total: " & pf.PivotItems.Count & vbLf & "Visible: " & pf.VisibleItems.Count
End Sub



Sub ShowVisible()
Dim ws As Worksheet, pt As PivotTable, pf As PivotField
Set ws = ActiveSheet
Set pt = ws.PivotTables("pivot_table5")
Set pf = pt.PivotFields("Familia")


For Each i In pf.PivotItems
If i.Visible = True Then
MsgBox i.Name
End If
Next i


End Sub
