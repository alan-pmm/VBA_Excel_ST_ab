VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' agregado desde aquí


Private Sub Worksheet_Calculate()

'TOOGLE BOX
'If pivot_sincro = True Then Exit Sub

'CELL
If Sheets("Ventas STD").Range("P7").Value = "pivot_desincro" Then
pivot_sincro_cell = False
Else
pivot_sincro_cell = True
End If


If pivot_sincro_cell = False Then
Exit Sub
End If

If ActiveWorkbook Is ThisWorkbook Then
Else
Exit Sub
End If

If ActiveWorkbook.Worksheets("Ventas EOY") Is ActiveSheet Then
Application.EnableEvents = False
Call CB_For_Each_Pi_select_val.main_macro_13
Application.EnableEvents = True
End If

End Sub

