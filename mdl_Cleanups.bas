Attribute VB_Name = "mdl_Cleanups"
Option Explicit


Private Sub HardcodeExternalReferences()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

'   Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

''' CODE GOES HERE
' Sets to original workbook

Dim ws As Worksheet
Dim wb As Workbook
Dim rng As Range
Dim cell As Range

Set wb = ActiveWorkbook

For Each ws In wb.Worksheets
    ws.Activate
        For Each cell In ws.UsedRange
            ' Check if the cell contains ".xls"
            If InStr(1, cell.Formula, ".xls") > 0 Then
                cell.Value = cell.Value
            End If
        Next cell
    ws.Cells(1, 1).Select
Next ws

wb.Sheets(1).Activate

Debug.Print "Complete"

'   Shutdown
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

Exit Sub

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub


