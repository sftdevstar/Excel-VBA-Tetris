VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPaused 
   Caption         =   "PAUSED"
   ClientHeight    =   840
   ClientLeft      =   30
   ClientTop       =   480
   ClientWidth     =   4785
   OleObjectBlob   =   "FrmPaused.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPaused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TheEnd     As Boolean

Private Sub NoBut_Click()
    TheEnd = True
    PausedForm.Hide
End Sub

Private Sub YesBut_Click()
    Sheet1.gameover = True
    TheEnd = True
    PausedForm.Hide
End Sub
