Attribute VB_Name = "mdl_Worksheet"
Option Explicit


Private Sub DisplayOutlines()
'shows the group/ungroup outline buttons if the default setting is changed for some reason
Dim ws As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook
For Each ws In wb.Worksheets
ws.Activate
Debug.Print ws.Name; ActiveWindow.DisplayOutline
ActiveWindow.DisplayOutline = True
Debug.Print ws.Name; ActiveWindow.DisplayOutline
Next ws

End Sub


Private Sub UnhideAllRows()

Dim ws As Worksheet
Dim screenUpdateState As Variant
Dim eventsState As Variant

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = ActiveSheet

ActiveSheet.Cells.EntireRow.Hidden = False

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState

End Sub


Public Sub HideRows(varRng As Range)
    Set varRng = varRng.SpecialCells(xlCellTypeBlanks)
    varRng.EntireRow.Hidden = True
End Sub

Private Sub HideEmptyRows()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim ws As Worksheet
Dim pt As PivotTable
Dim maxRow As Variant
Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden

Dim varRng As Range

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = ActiveSheet

maxRow = ws.UsedRange.Rows.Count

'   Initial area to be hidden
Set varRng = ws.Range("A10").Resize(maxRow)

Call HideRows(varRng)

For Each pt In ws.PivotTables
    pt.TableRange1.EntireRow.Hidden = False                 ' Unhide pivot area that was mistakenly hidden

    Set varRng = pt.DataLabelRange
    varRng.Offset(-x).Resize(x).EntireRow.Hidden = False    ' Unhide buffer area above pivottable
Next pt

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
End Sub




Private Sub HideEmptyRowsA()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim ws As Worksheet
Dim maxRow As Variant
Dim i As Variant
Dim j As Variant
Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = ActiveSheet

'ws.UsedRange ' Refresh UsedRange
maxRow = ws.UsedRange.Rows.Count

For i = 15 To maxRow
    If ActiveSheet.Cells(i, 1) = "" Then
        ActiveSheet.Cells(i, 1).EntireRow.Hidden = True
    Else
        For j = i To i - x Step -1
            ActiveSheet.Cells(j, 1).EntireRow.Hidden = False
        Next j
    End If
Next i

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
End Sub

Private Sub AutofitCol()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim ws As Worksheet
Dim maxCol As Variant
Dim i As Variant
Dim x As Variant: x = 2         ' Additional column width added to autofit

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = ActiveSheet

'ws.UsedRange ' Refresh UsedRange
maxCol = ws.UsedRange.Columns.Count

For i = 1 To maxCol
    ws.Columns(i).EntireColumn.AutoFit
    ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth + x
Next i

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
End Sub

Private Function GetColumnLetter(lngCol As Long) As String

Dim vArr
vArr = Split(Cells(1, lngCol).Address(True, False), "$")
GetColumnLetter = vArr(0)

End Function

''' GET UNDO LIST

' For i = 1 to Application.CommandBars("Standard").Controls("&Undo").Control.ListCount
'   Debug.Print Application.CommandBars("Standard").Controls("&Undo").Control.List(i)
' Next i

''' HEADINGS

Private Sub hideHeadings()

Dim ws As Worksheet
Dim currws As Worksheet

    On Error GoTo errHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

Set currws = ActiveSheet

For Each ws In Worksheets
    ws.Activate
    ActiveWindow.DisplayHeadings = False
Next ws

currws.Activate

    Application.ScreenUpdating = True
    Application.EnableEvents = True
errHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Private Sub showHeadings()

Dim ws As Worksheet
Dim currws As Worksheet

    On Error GoTo errHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

Set currws = ActiveSheet

For Each ws In Worksheets

    ws.Activate
    ActiveWindow.DisplayHeadings = True

Next ws

currws.Activate

    Application.ScreenUpdating = True
    Application.EnableEvents = True
errHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


'''''''''''''' TAB RENAME

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    Set Target = Range("A2")
    If Target = "" Then Exit Sub
    On Error GoTo Badname
    ActiveSheet.Name = Left(Target, 31)
    Exit Sub
Badname:
    MsgBox "Please revise the entry in " & Target.AddressLocal & Chr(13) _
    & "It appears to contain one or more " & Chr(13) _
    & "illegal characters or is a duplicate." & Chr(13)
    Target.Activate
End Sub


Option Explicit

'Sub FindingLastRow_Alternatives()
'
'Dim sht As Worksheet
'Dim LastColumn As Long
'
'Set sht = ThisWorkbook.Worksheets(Sheet1.Name)
'
''Provided by Bob U.
'  LastRow = sht.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
'
'End Sub



'Sub ReportRowCol()
'    Call UnhideAllRows
'    Call HideEmptyRows
'    Call AutofitCol
'End Sub
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

''''    NEW METHOD TO HIDE ROWS (both subs below

'Public Sub HideRows(varRng As Range)
'    Set varRng = varRng.SpecialCells(xlCellTypeBlanks)
'    varRng.EntireRow.Hidden = True
'End Sub
'
'Sub HideEmptyRows()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim maxRow As Variant
'Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden
'
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
'For Each pt In ws.PivotTables
'    pt.TableRange1.EntireRow.Hidden = False                 ' Unhide pivot area that was mistakenly hidden
'
'    Set varRng = pt.DataLabelRange
'    varRng.Offset(-x).Resize(x).EntireRow.Hidden = False    ' Unhide buffer area above pivottable
'Next pt
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub



'''' OLD Method to hide rows, loop through. slow method

'Sub HideEmptyRows()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim i As Variant
'Dim j As Variant
'Dim x As Variant: x = 5         ' Buffer rows above pivots that are not hidden
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'For i = 15 To 500
'    If ActiveSheet.Cells(i, 1) = "" Then
'        ActiveSheet.Cells(i, 1).EntireRow.Hidden = True
'    Else
'        For j = i To i - x Step -1
'            ActiveSheet.Cells(j, 1).EntireRow.Hidden = False
'        Next j
'    End If
'Next i
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub


''''''''''  AUTOFIT COLUMNS w/ buffer

'Sub AutofitCol()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim i As Variant
'Dim x As Variant: x = 2         ' Additional column width added to autofit
'
'On Error GoTo errHandler
'screenUpdateState = Application.Scree nUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Set ws = ActiveSheet
'
'For i = 1 To 100
'    ws.Columns(i).EntireColumn.AutoFit
'        Debug.Print ws.Columns(i).ColumnWidth
'    ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth + x
'        Debug.Print ws.Columns(i).ColumnWidth
'Next i
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub


