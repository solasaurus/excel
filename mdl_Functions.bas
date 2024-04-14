Attribute VB_Name = "mdl_Functions"
Option Explicit


'
'Function ConcatenateIf(CriteriaRange As Range, Condition As Variant, ConcatenateRange As Range, Optional Separator As String = ",") As Variant
''Update 20150414
'Dim xResult As String
'On Error Resume Next
'If CriteriaRange.Count <> ConcatenateRange.Count Then
'    ConcatenateIf = CVErr(xlErrRef)
'    Exit Function
'End If
'For i = 1 To CriteriaRange.Count
'    If CriteriaRange.Cells(i).Value = Condition Then
'        xResult = xResult & Separator & ConcatenateRange.Cells(i).Value
'    End If
'Next i
'If xResult <> "" Then
'    xResult = VBA.Mid(xResult, VBA.Len(Separator) + 1)
'End If
'ConcatenateIf = xResult
'Exit Function
'End Function

'
'Function ColorFunction(rColor As Range, rRange As Range, Optional SUM As Boolean)
'Dim rCell As Range
'Dim lCol As Long
'Dim vResult
'lCol = rColor.Interior.ColorIndex
'If SUM = True Then
'For Each rCell In rRange
'If rCell.Interior.ColorIndex = lCol Then
'vResult = WorksheetFunction.SUM(rCell, vResult)
'End If
'Next rCell
'Else
'For Each rCell In rRange
'If rCell.Interior.ColorIndex = lCol Then
'vResult = 1 + vResult
'End If
'Next rCell
'End If
'ColorFunction = vResult
'End Function
