Attribute VB_Name = "Module1"
Const PW = "tet_123%"
Sub Protect_Proj()
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Sheets
        sht.Protect _
            password:=PW, _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True, _
            UserInterfaceOnly:=True
    Next sht
End Sub



Sub test4()
    Dim sht As Worksheet
    Dim rng As Range
    
    Set sht = Worksheets("TETRIS")
    
    For Each rng In sht.Range("quadro").Rows
        rng.RowHeight = 16
    Next rng
    
    For Each rng In sht.Range("quadro").Columns
        rng.ColumnWidth = 2.3
    Next rng
  
End Sub


Sub test5()
    Dim sht As Worksheet
    Dim rng As Range
    
    Set sht = Worksheets("TETRIS")
    
    Set rng = sht.Range("quadro")
'    Debug.Print sht.Range("quadro").Rows(1).Height
'    Debug.Print sht.Range("quadro").Columns(1).Width

    Debug.Print rng.Row
    Debug.Print rng.Column
    Debug.Print rng.Rows.count
    Debug.Print rng.Columns.count
 
End Sub


Sub test6()
    Dim sht As Worksheet
    Dim rng As Range
    Dim i As Integer
    Set sht = Worksheets("TETRIS")
    sht.Columns(1).ColumnWidth = 2.3
    sht.Rows(25).RowHeight = 16
    For i = 1 To 10
        Set rng = sht.Columns(12 + i)
        rng.ColumnWidth = 2.3
    Next i
End Sub


Sub test7()
    Dim sht As Worksheet
    Dim shp As Shape

    Set sht = Worksheets("TETRIS")
    
    For Each shp In sht.Shapes
        If InStr(1, shp.Name, "figt") > 0 Or _
           InStr(1, shp.Name, "figs") > 0 Or _
           InStr(1, shp.Name, "fig") > 0 Then
            shp.Visible = msoFalse
        Else
            shp.Visible = msoTrue
        End If
    Next shp
End Sub

