Attribute VB_Name = "mdl_Tables"
Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''
'       FILTERS

Sub tbl_FilterReset(tbl As ListObject)

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

Dim AutoFilterState As Boolean

''' Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

With tbl
    AutoFilterState = .ShowAutoFilter
    .Range.AutoFilter
    .ShowAutoFilter = AutoFilterState
End With

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub

Private Sub TableFilterOff()
Dim ws As Worksheet
Dim tbl As ListObject
Set ws = Worksheets("")
Set tbl = ws.ListObjects("")
Call tbl_FilterButtonOff(tbl)
End Sub

Private Sub tbl_FilterButtonOff(tbl As ListObject)

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

''' Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

With tbl
    .Range.AutoFilter
    .ShowAutoFilter = False
End With

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub

Private Sub tbl_HideFilterDropdowns_ByColumn()
Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

Dim ws As Worksheet
Dim tbl As ListObject
Dim arr1(1 To 8)
Dim arr2(1 To 8)
Dim c As Variant
Dim i As Variant

''' Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Set ws = Sheets("")
Set tbl = ws.ListObjects("")

'Load Columns that dropdowns will be removed from
arr1(1) = "COL 1"
arr1(2) = "COL 2"
arr1(3) = "COL 3"

'Load the index of the column name
For i = LBound(arr1) To UBound(arr1)
    arr2(i) = tbl.ListColumns(arr1(i)).Index
Next i

'Loops through columns, if column name = column in array1 then column (index number / array2) dropdown is removed
For Each c In tbl.ListColumns
    For i = LBound(arr1) To UBound(arr1)
        If arr1(i) = c Then
            tbl.Range.AutoFilter field:=arr2(i), VisibleDropDown:=False
        End If
    Next i
Next c

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub

Private Sub tbl_Tracker_AutoFilters()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim ws As Worksheet
Dim tbl As ListObject
Dim strVar As String

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = Sheets("")
Set tbl = ws.ListObjects("")

' Calculate before filtering on columns that have calculated fields
ws.Calculate

With tbl
    strVar = "=" & ws.OLEObjects("cbx_ProjectName").Object.Value
    If Not strVar = "=" Then
        .Range.AutoFilter field:=.ListColumns("PROJECT NAME").Index, _
            Criteria1:=strVar
    End If
End With

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''
'       SORTS

Private Sub tbl_Closings_Sort(tbl As ListObject)

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant

''' Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

With tbl.Sort
    .SortFields.Clear
    .SortFields.Add Key:=tbl.ListColumns("COLUMN NAME 1").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=tbl.ListColumns("COLUMN NAME 2").Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .SortFields.Add Key:=tbl.ListColumns("COLUMN NAME 3").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=tbl.ListColumns("COLUMN NAME 4").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Header = xlYes
    .Apply
    .SortFields.Clear
End With

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState
Application.Calculation = calcState

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
'       DATA DELETE


Private Sub DeleteAllTableData()

Dim ws As Worksheet
Dim tbl As ListObject

Set ws = Sheets("")
Set tbl = ws.ListObjects("")

Call tbl_DeleteAllData(tbl)

End Sub

Private Sub tbl_DeleteAllData(tbl As ListObject)

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim tRows As Long
Dim tCols As Long
Dim r As Long
Dim c As Long

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Call tbl_FilterReset(tbl)

With tbl.DataBodyRange
    tRows = .Rows.Count
    tCols = .Columns.Count
End With

For r = tRows To 1 Step -1
    tbl.ListRows(r).Delete
Next r

tbl.ListRows.Add (1)

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
'       ROW / COLUMN IDENTIFICATION


'Finds the number of the column in a table, given the name/string of the column name
Private Function ColLookup(ByVal columnname As String)
    ColLookup = Evaluate("=MATCH(" & Chr(34) & columnname & Chr(34) & ",(PIPELINE[#Headers]),0)")
End Function


'Finds the letter of the column address, given a the column number/integer
Private Function GetColumnLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function



