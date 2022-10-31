VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Type TpPos
    r          As Integer
    c          As Integer
End Type
Private Type tpfig
    count      As Integer
    cor        As Integer
    stop       As Boolean
    pos(1 To 4) As TpPos
    nextpos(1 To 4) As TpPos
    rot        As Integer
    rotdata(1 To 4, 1 To 4) As TpPos
End Type
Dim fig        As tpfig
Dim conjfig(1 To 10) As tpfig
Dim copo(0 To 25, -1 To 14) As Boolean
Public gameover As Boolean
Dim speed      As Single
Dim prox_fig   As Integer

Private Sub MessageBack(ByVal s As String, ti As Integer, x As Integer, y As Integer)
    Dim ts     As String
    If ti <= 9 Then
        ts = "0" & ti
    Else
        ts = ti
    End If
    With Shapes("game")
        .TextEffect.Text = s
        .ZOrder (msoBringToFront)
        .Top = y
        .Left = x
        .Visible = True
    End With
    Application.OnTime Now + TimeValue("00:00:" & ts), "sheet1.DeleteMsg"
End Sub

Public Sub DeleteMsg()
    With Shapes("game")
        .Top = 1000
        .Left = 1000
        .Visible = False
    End With
    
End Sub

Function Vscore(s As String) As Long
    Vscore = CLng(Shapes(s).TextEffect.Text)
End Function

Sub AddScore(s As Long)
    Dim p      As Long
    p = CLng(Shapes("score").TextEffect.Text)
    p = p + s
    Shapes("score").TextEffect.Text = Format(p, "000000")
End Sub

Private Sub Proc()
    Dim r      As Integer, c As Integer
    Dim i      As Integer, x As Integer, k As Integer, fundo As Boolean
    Dim vazio  As Boolean, mudou As Boolean, v As Integer, rx As Integer
    r = Range("quadro").Rows.count
    c = Range("quadro").Columns.count
    mudou = False
    Sheet1.Unprotect ("tet_123%")
    v = 0
    Do While r > 3
        i = 1
        vazio = False
        Do While i <= c And Not vazio
            vazio = Not copo(r, i)
            i = i + 1
        Loop
        If Not vazio Then
            If Not mudou Then
                mudou = True
                'Application.ScreenUpdating = False
            End If
            AddScore (3 * r)
            v = v + 1
            rx = Range("quadro").Cells(r, 1).Top
            For x = r To 2 Step -1
                For k = 1 To c
                    Shapes("figt" & FP(x) & FP(k)).Visible = False
                    copo(x, k) = False
                    If copo(x - 1, k) Then
                        Shapes("figt" & FP(x) & FP(k)).Fill.ForeColor.RGB = Shapes("figt" & FP(x - 1) & FP(k)).Fill.ForeColor.RGB
                        Shapes("figt" & FP(x) & FP(k)).Visible = True
                        copo(x, k) = True
                    End If
                Next
            Next
        Else
            r = r - 1
        End If
    Loop
    'If mudou Then Application.ScreenUpdating = True
    If v > 2 Then
        AddScore (50 * v)
        DoEvents
        MessageBack 50 * v, 1, Range("quadro").Cells(rx, 5).Left, Range("quadro").Cells(13, 1).Top
    End If
    Call Protect_Proj
End Sub

Private Sub Show_Next(fig As tpfig)
    Dim i      As Integer, rgbcor As Long
    Range("next").Clear
    For i = 1 To 4
        Shapes("figs" & i).Visible = False
    Next
    With fig
        Range("next").Cells(.pos(1).r, .pos(1).c - 4).Interior.ColorIndex = .cor
        rgbcor = Range("next").Cells(.pos(1).r, .pos(1).c - 4).Interior.Color
        For i = 1 To .count
            Shapes("figs" & i).Fill.ForeColor.RGB = rgbcor
            Shapes("figs" & i).Top = Range("next").Cells(.pos(i).r, .pos(i).c - 4).Top
            Shapes("figs" & i).Left = Range("next").Cells(.pos(i).r, .pos(i).c - 4).Left
            Shapes("figs" & i).Visible = True
        Next
    End With
End Sub

Private Sub Config(fig As tpfig)
    Dim i      As Integer, j As Integer, k As Integer
    Cells(21, 20).Activate
    With conjfig(1)
        .count = 4
        .cor = 40
    End With
    With conjfig(2)
        .count = 1
        .cor = 6
    End With
    With conjfig(3)
        .count = 2
        .cor = 46
    End With
    With conjfig(4)
        .count = 3
        .cor = 33
    End With
    With conjfig(5)
        .count = 4
        .cor = 4
    End With
    With conjfig(6)
        .count = 4
        .cor = 7
    End With
    With conjfig(7)
        .count = 4
        .cor = 41
    End With
    With conjfig(8)
        .count = 4
        .cor = 50
    End With
    With conjfig(9)
        .count = 4
        .cor = 54
    End With
    With conjfig(10)
        .count = 4
        .cor = 15
    End With
    Randomize
    For k = 1 To 10
        With conjfig(k)
            For i = 1 To 4
                .pos(i).r = Range("ROT_DATA").Cells(1 + (k - 1) * 5, i).Value
                .pos(i).c = Range("ROT_DATA").Cells(1 + (k - 1) * 5, i + 4).Value
            Next
            For i = 1 To 4
                For j = 1 To 4
                    .rotdata(i, j).r = Range("ROT_DATA").Cells(i + 1 + (k - 1) * 5, j).Value
                    .rotdata(i, j).c = Range("ROT_DATA").Cells(i + 1 + (k - 1) * 5, j + 4).Value
                Next
            Next
            .stop = False
            .rot = 1
        End With
    Next
    If prox_fig = 0 Then
        fig = conjfig(Int(1 + 10 * Rnd))
        Call Show_Next(fig)
        prox_fig = (Int(1 + 10 * Rnd))
    Else
        fig = conjfig(prox_fig)
        prox_fig = (Int(1 + 10 * Rnd))
    End If
    For i = 1 To 4
        Shapes("fig" & i).Visible = False
        Shapes("fig" & i).Fill.ForeColor.RGB = Shapes("figs" & i).Fill.ForeColor.RGB
    Next
    For i = 1 To 23
        For j = 1 To 12
            If Not copo(i, j) Then
                Shapes("figT" & FP(i) & FP(j)).Fill.ForeColor.RGB = Shapes("fig1").Fill.ForeColor.RGB
            End If
        Next
    Next
    Call Show_Next(conjfig(prox_fig))
    With fig
        .nextpos(1) = .pos(1)
        .nextpos(2) = .pos(2)
        .nextpos(3) = .pos(3)
        .nextpos(4) = .pos(4)
    End With
End Sub

Function FP(ByVal s As String) As String
    If Len(s) = 1 Then FP = "0" & s Else FP = s
End Function

Private Function TheEnd(fig As tpfig) As Boolean
    Dim i      As Integer
    TheEnd = False
    With fig
        For i = 1 To .count
            If .pos(i).r < 3 Then
                TheEnd = True
            End If
        Next
    End With
End Function

Private Sub MoveDown(fig As tpfig)
    Dim i      As Integer, r As Integer, c As Integer
    With fig
        For i = 1 To .count
            .nextpos(i).r = .pos(i).r + 1
            .nextpos(i).c = .pos(i).c
        Next
    End With
    If IsOkToMove(fig) Then
        Call Draw(fig)
    Else
        With fig
            .stop = True
            For i = 1 To .count
                copo(.pos(i).r, .pos(i).c) = True
            Next
            AddScore (.pos(1).r)
        End With
        Call Proc
        If TheEnd(fig) Then
            gameover = True
            Shapes("gameover").Visible = True
            Shapes("gameover").Top = Cells(15, 1).Top
        End If
    End If
End Sub

Private Function IsOkToMove(fig As tpfig) As Boolean
    Dim ok     As Boolean
    Dim i      As Integer
    ok = True
    With fig
        For i = 1 To .count
            If copo(.nextpos(i).r, .nextpos(i).c) Then
                ok = False
            End If
        Next
    End With
    IsOkToMove = ok
End Function

Private Sub Draw(fig As tpfig)
    Dim i      As Integer
    With fig
        For i = 1 To .count
            Shapes("figt" & FP(.pos(i).r) & FP(.pos(i).c)).Visible = False
            .pos(i) = .nextpos(i)
        Next
        For i = 1 To .count
            Shapes("figt" & FP(.pos(i).r) & FP(.pos(i).c)).Visible = True
        Next
    End With
End Sub

Private Sub CommandButton1_Click()
    FrmPaused.Show
End Sub

Sub About()
    FrmAbout.Show
End Sub

Sub Play()
    If Not Shapes("play").Visible Then
        FrmPaused.Show
        If gameover Then
            Shapes("play").Visible = True
            Shapes("stop").Visible = False
        End If
    Else
        Call Play_Now
    End If
End Sub

Private Sub Play_Now()
    Dim t      As Double
    Dim sp     As Single
    Dim score  As Long, nome As String
    Dim i      As Integer, j As Integer
    Dim saved(1 To 4) As TpPos
    Sheet1.ScrollArea = "A1:Y27"
    Call Protect_Proj
    Shapes("gameover").Visible = False
    Shapes("gameover").Top = Cells(27, 1).Top
    Range("quadro").Value = 0
    Shapes("play").Visible = False
    Shapes("stop").Visible = True
    Application.Cursor = xlNorthwestArrow
    Shapes("score").TextEffect.Text = "000000"
    Application.StatusBar = ""
    Application.EnableEvents = True
    gameover = False
    Range("quadro").Clear
    For i = 0 To 25
        For j = -1 To 14
            copo(i, j) = True
        Next
    Next
    For i = 1 To 23
        For j = 1 To 12
            copo(i, j) = False
            nome = "figt" & FP(i) & FP(j)
            With Shapes(nome)
                .Visible = False
                .Top = Range("quadro").Cells(i, j).Top
                .Left = Range("quadro").Cells(i, j).Left
            End With
        Next
    Next
    Cells(21, 20).Activate
    Call Config(fig)
    Call Draw(fig)
    sp = 0.52
    speed = sp
    On Error GoTo fim
ini:
    Do While Not gameover
        For i = 1 To fig.count
            saved(i) = fig.pos(i)
        Next
        Call MoveDown(fig)
        t = Timer
        Do While Timer - t < speed
            DoEvents
        Loop
        If fig.stop Then
            Call Config(fig)
            score = Vscore("score")
            If score > 30000 Then
                sp = 0.06
            Else
                If score > 20000 Then
                    sp = 0.1
                Else
                    If score > 10000 Then
                        sp = 0.14
                    Else
                        If score > 5000 Then
                            sp = 0.16
                        End If
                    End If
                End If
            End If
            speed = sp
        End If
    Loop
    If Vscore("score") > Vscore("hiscore") Then
        Shapes("hiscore").TextEffect.Text = Shapes("score").TextEffect.Text
    End If
    Application.StatusBar = False
    Shapes("play").Visible = True
    Shapes("stop").Visible = False
    gameover = True
    Application.Cursor = xlDefault
    'Sheet1.Protect
    Exit Sub
fim:
    For i = 1 To fig.count
        fig.pos(i) = saved(i)
    Next
    Resume Next
    
End Sub

Private Sub PlayBut_Click()
    Call Play
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r      As Integer, c As Integer
    Dim i      As Integer
    Application.EnableEvents = False
    If Not gameover And Not fig.stop Then
        r = Target.Row
        c = Target.Column
        With fig
            If c = 21 Then
                For i = 1 To .count
                    .nextpos(i).r = .pos(i).r
                    .nextpos(i).c = 1 + .pos(i).c
                Next
                If IsOkToMove(fig) Then
                    Call Draw(fig)
                End If
            Else
                If c = 19 Then
                    For i = 1 To .count
                        .nextpos(i).r = .pos(i).r
                        .nextpos(i).c = .pos(i).c - 1
                    Next
                    If IsOkToMove(fig) Then
                        Call Draw(fig)
                    End If
                Else
                    If r = 20 Then
                        For i = 1 To .count
                            .nextpos(i).r = .pos(i).r + .rotdata(.rot, i).r
                            .nextpos(i).c = .pos(i).c + .rotdata(.rot, i).c
                        Next
                        If IsOkToMove(fig) Then
                            .rot = .rot + 1
                            If .rot = 5 Then .rot = 1
                            Call Draw(fig)
                        End If
                    Else
                        If r = 22 Then
                            speed = 0.02
                        End If
                    End If
                End If
            End If
            Cells(21, 20).Activate
        End With
    End If
    Application.EnableEvents = True
End Sub
