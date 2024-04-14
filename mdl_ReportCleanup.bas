Attribute VB_Name = "mdl_ReportCleanup"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''
''' AGGREGATE SUBS
'
'Sub ReportCleanup_pvt()
'    Call UnhideAllRows
'    Call HideEmptyRows_pvt
'    Call AutofitCol
'End Sub
'
'Sub ReportCleanup_tbl()
'    Call UnhideAllRows
'    Call HideEmptyRows_tbl
'    Call AutofitCol
'End Sub


''''''''''''''''''''''''''''''''''''''''''''''''
'''' PIVOT & TABLE SUBS
'
'Sub UnhideAllRows()
'
'Dim ws As Worksheet
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Set ws = ActiveSheet
'
'ActiveSheet.Cells.EntireRow.Hidden = False
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'
'
'Public Sub HideRows(varRng As Range)
'    Set varRng = varRng.SpecialCells(xlCellTypeBlanks)
'    varRng.EntireRow.Hidden = True
'End Sub
'
'Sub AutofitCol()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim maxCol As Variant
'Dim i As Variant
'Dim x As Variant: x = 2         ' Additional column width added to autofit
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Set ws = ActiveSheet
'
''ws.UsedRange ' Refresh UsedRange
'maxCol = ws.UsedRange.Columns.Count
'
'For i = 1 To maxCol
'    ws.Columns(i).EntireColumn.AutoFit
'    ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth + x
'Next i
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub
'
''''''''''''''''''''''''''''''''''''''''''''''''
'''' PIVOT EXCLUSIVE SUBS
'
'

'
'
''Sub HideEmptyRowsA()
''
''Dim screenUpdateState As Variant
''Dim eventsState As Variant
''Dim ws As Worksheet
''Dim maxRow As Variant
''Dim i As Variant
''Dim j As Variant
''Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden
''
''On Error GoTo errHandler
''screenUpdateState = Application.ScreenUpdating
''eventsState = Application.EnableEvents
''Application.ScreenUpdating = False
''Application.EnableEvents = False
''
''Set ws = ActiveSheet
''
'''ws.UsedRange ' Refresh UsedRange
''maxRow = ws.UsedRange.Rows.Count
''
''For i = 15 To maxRow
''    If ActiveSheet.Cells(i, 1) = "" Then
''        ActiveSheet.Cells(i, 1).EntireRow.Hidden = True
''    Else
''        For j = i To i - x Step -1
''            ActiveSheet.Cells(j, 1).EntireRow.Hidden = False
''        Next j
''    End If
''Next i
''
''errHandler:
''Application.ScreenUpdating = screenUpdateState
''Application.EnableEvents = eventsState
''End Sub
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''
'''' TABLE EXCLUSIVE SUBS
'
'
'Sub HideEmptyRows_tbl()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim tbl As ListObject
'Dim maxRow As Variant
'Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden
'Dim varRng As Range
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Set ws = ActiveSheet
'
'maxRow = ws.UsedRange.Rows.Count
'
''   Initial area to be hidden
'Set varRng = ws.Range("A10").Resize(maxRow)
'
'Call HideRows(varRng)
'
'For Each tbl In ws.ListObjects
'    Set varRng = tbl.HeaderRowRange                                     '   Set range to table header row
'    varRng.Offset(-x).Resize(x).EntireRow.Hidden = False                '   Unhides buffer area above table
'    varRng.Resize(tbl.ListRows.Count + 1).EntireRow.Hidden = False      '   Unhides any blank table rows that were hidden
'
'    Set varRng = tbl.ListRows(tbl.ListRows.Count + 1).Range             '   Set range to row after table
'    varRng.Resize(x).EntireRow.Hidden = False                           '   Unhides buffer area below table
'Next tbl
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub
'
'
