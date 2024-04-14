Attribute VB_Name = "mdl_Shapes"
Option Explicit

Private Sub ShowHideShapes()

    For i = 1 To ActiveSheet.Shapes.Count
    ActiveSheet.Shapes(i).Visible = msoFalse
    Next i
    For i = 1 To ActiveSheet.Shapes.Count
    ActiveSheet.Shapes(i).Visible = msoTrue
    Next i

End Sub

'
'Sub ShowShapes()
'    For i = 1 To ActiveSheet.Shapes.Count
'    ActiveSheet.Shapes(i).Visible = True
'    Next i
'End Sub
'
'
'Sub HideShapes()
'    For i = 1 To ActiveSheet.Shapes.Count
'    ActiveSheet.Shapes(i).Visible = False
'    Next i
'End Sub
'
'' Flips shape visibility to opposite of current setting
'Sub ShowHideShapes()
'    For i = 1 To ActiveSheet.Shapes.Count
'    ActiveSheet.Shapes(i).Visible = IIf(ActiveSheet.Shapes(i).Visible = True, False, True)
'    Next i
'End Sub
'
'
'' Toggle shape visibility of array with known shape names
'Sub ShapeViewToggle()
'
'Dim ws As Worksheet
'Dim arrShapes As Variant
'Dim i As Integer
'Dim shpVar As Shape
'
'Set ws = Sheets("WorksheetName")
'
'arrShapes = Array( _
'                "Shape1", _
'                "Shape2", _
'                "Shape3" _
'                )
'
'For i = 0 To UBound(arrShapes, 1)
'    On Error GoTo NextLoop1
'    Set shpVar = ws.Shapes(arrShapes(i))
'    shpVar.Visible = IIf(shpVar.Visible = True, False, True)
'    Set shpVar = Nothing
'NextLoop1:
'Next i
'
'End Sub

