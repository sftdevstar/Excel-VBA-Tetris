VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Sheet1.Shapes("play").Visible = True
    Sheet1.Shapes("stop").Visible = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call Protect_Proj
End Sub
