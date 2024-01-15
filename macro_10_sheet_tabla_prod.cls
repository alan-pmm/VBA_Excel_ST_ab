VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Call copyLastfilVAL3
End Sub

Private Sub Worksheet_Calculate()
Application.EnableEvents = False
Call copyLastfilVAL3_macro_10
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Call copyfirstfilVAL2
End Sub





