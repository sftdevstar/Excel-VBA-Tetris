VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAbout 
   Caption         =   "About"
   ClientHeight    =   1350
   ClientLeft      =   30
   ClientTop       =   480
   ClientWidth     =   3780
   OleObjectBlob   =   "FrmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TheEnd     As Boolean
Private Sub CommandButton1_Click()
    TheEnd = True
    UserForm1.Hide
End Sub

Private Sub UserForm_Activate()
    Dim t      As Single
    TheEnd = False
    With UserForm1
        Do
            .Image1.Left = .Image1.Left + 3
            If .Image1.Left > CommandButton1.Left - 10 Then .Image1.Left = 10
            t = Timer
            Do
                DoEvents
            Loop Until Timer - t > 0.1
        Loop Until TheEnd
    End With
End Sub
