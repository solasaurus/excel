Attribute VB_Name = "mdl_Pivots"
Option Explicit


Sub pvt_ValueFormat()

    Dim pt As PivotTable
    Dim pf As PivotField
    
    On Error Resume Next
    Set pt = ActiveCell.PivotCell.PivotTable
    If Err.Number <> 0 Then
        MsgBox "The cursor needs to be in a pivot table"
        Exit Sub
    End If
    
    'For Each pt In ActiveSheet.PivotTables
        For Each pf In pt.DataFields
            'Debug.Print pf.Name
            pf.Function = xlSum
            pf.NumberFormat = "#,##0"
        Next pf
    'Next pt
    
End Sub


Private Sub pvt_Refresh()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim ws As Worksheet
Dim pf As PivotField
Dim pt As PivotTable

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
Application.ScreenUpdating = False
Application.EnableEvents = False

Set ws = ActiveSheet

For Each pt In ws.PivotTables
    pt.RefreshTable
    '   The following prevents pivot items from being stored in the pivot cache
    '       Otherwise when filtering pivots via vba, errors result when a pi loop lands on a old/nonexistant pivot item
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
Next pt

errHandler:
Application.ScreenUpdating = screenUpdateState
Application.EnableEvents = eventsState

End Sub


Private Sub PrintPivotFilterItems()
Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant
Dim ws As Worksheet
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem

Dim strFilter As String

Set ws = ActiveSheet

On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'strFilter = "Account: Account Number"

For Each pt In ws.PivotTables
Debug.Print ws.Name & " " & pt.Name
    For Each pf In pt.PivotFields
        If pf.Orientation = xlPageField Then
        'If pf.Name = strFilter Then
            Debug.Print pf.Name
            For Each pi In pf.PivotItems
                If pi.Visible = True Then
                    Debug.Print pi.Name
                End If
            Next pi
        End If
    Next pf
Next pt

errHandler:
    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState

End Sub

'Sub HideEmptyRows_pvt()
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
'On Error Resume Next
'
'For Each pt In ws.PivotTables
'    pt.TableRange1.EntireRow.Hidden = False                 '   Unhide pivot area that was mistakenly hidden
'
'    Set varRng = pt.DataLabelRange
'    varRng.Offset(-x).Resize(x).EntireRow.Hidden = False    '   Unhide buffer area above pivottable
'
'    pt.PageRange.EntireRow.Hidden = True                    '   Rehide Pivot Filters
'
'Next pt
'
'On Error GoTo 0
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'End Sub

'Sub Pvt_ClearFilters()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Dim ws As Worksheet
'Dim pf As PivotField
'Dim pt As PivotTable
'
'Set ws = ActiveSheet
'
''For Each ws In ThisWorkbook.Worksheets
'For Each pt In ws.PivotTables
'    pt.ClearAllFilters
'Next pt
''Next ws
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub

'''   BASIC REFRESH PIVOT TABLES
'Sub pvtRefresh()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim pf As PivotField
'Dim pt As PivotTable
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Set ws = ActiveSheet
'
''For Each ws In ThisWorkbook.Worksheets
'For Each pt In ws.PivotTables
'    pt.RefreshTable
'Next pt
''Next ws
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'

''' REFRESH PIVOT AND CLEAR PIVOT CACHE (sets store cache to zero)
'Sub pvtRefresh()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim pf As PivotField
'Dim pt As PivotTable
'Dim CurrLockMode As Integer
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'
'Set ws = ActiveSheet
'
'For Each pt In ws.PivotTables
'    pt.RefreshTable
'    '   The following prevents pivot items from being stored in the pivot cache
'    '       Otherwise when filtering pivots via vba, errors result when a pivot item loop lands on a old/nonexistant pivot item
'    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
'Next pt
'
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub



'Sub FilterPivotsCustomValue()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim calcState As Variant
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim pf As PivotField
'Dim pi As PivotItem
'
'Dim strField1 As String
'Dim varValue1 As Variant
'
'Dim strField2 As String
'Dim varValue2 As Variant
'
'Dim strField3 As String
'Dim varValue3 As Variant
'
'Set ws = ActiveSheet
'
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'calcState = Application.Calculation
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'Application.Calculation = xlCalculationManual
'
'
'strField1 = "State"
'varValue1 = "MO"
'
'strField2 = "Credit Type"
'varValue2 = "LIHTC"
'
'strField3 = "Credit Year"
'varValue3 = "2017"
'
'
'For Each pt In ws.PivotTables
'    For Each pf In pt.PivotFields
'        Debug.Print pt.name & " " & pf.name
'        Select Case pf.name
'            Case strField1
'                pf.ClearAllFilters
'                For Each pi In pf.PivotItems
'                    pi.Visible = IIf(pi.Value = varValue1, True, False)
''                    If pi.Value = varValue1 Then
''                        pi.Visible = True
''                        'Exit For
''                    Else
''                        pi.Visible = False
''                    End If
'                Next pi
'            Case strField2
'                pf.ClearAllFilters
'                For Each pi In pf.PivotItems
'                    pi.Visible = IIf(pi.Value = varValue2, True, False)
''                    If pi.Value = varValue2 Then
''                        pi.Visible = True
''                        'Exit For
''                    Else
''                        pi.Visible = False
''                    End If
'                Next pi
'            Case strField3
'                pf.ClearAllFilters
'                For Each pi In pf.PivotItems
'                    pi.Visible = IIf(Val(pi.Value) >= Val(varValue3), True, False)
''                    If Val(pi.Value) >= Val(varValue3) Then
''                        pi.Visible = True
''                    Else
''                        pi.Visible = False
''                    End If
'                Next pi
'        End Select
'    Next pf
'Next pt
'
'errHandler:
'    Application.ScreenUpdating = screenUpdateState
'    Application.EnableEvents = eventsState
'    Application.Calculation = calcState
'
'End Sub



''==================================================================================================
''   DESCRIPTION:    Filters all pivot tables in the worksheet based on the inputs placed into a
''                   structured table in the worksheet. Must put a table at top of worksheet with
''                   "Field", "Value", & "Operator" columns
''   DATE CREATED:   2/23/2017
''   DATE UPDATED:   2/23/2017
''==================================================================================================
'Sub FilterAllPivotTables()
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim calcState As Variant
'Dim ws As Worksheet
'Dim tbl As ListObject
'Dim pt As PivotTable
'Dim pf As PivotField
'Dim pi As PivotItem
'Dim i As Integer
'Dim r As Integer
'
'Dim strField1 As String
'Dim varValue1 As Variant
'
'Dim strField2 As String
'Dim varValue2 As Variant
'
'Dim strField3 As String
'Dim varValue3 As Variant
'
'Dim arrField() As Variant
'Dim arrValue() As Variant
'Dim arrOperator() As Variant
'
'Dim strProcedure As String
'Dim strVar As String
'
'On Error GoTo errHandler
'    screenUpdateState = Application.ScreenUpdating
'    eventsState = Application.EnableEvents
'    calcState = Application.Calculation
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'
'strProcedure = "Filtering Pivots"
'
'Set ws = ActiveSheet
'Set tbl = ws.ListObjects(1)
'
'r = tbl.ListRows.Count
'
'ReDim arrField(1 To r)
'ReDim arrValue(1 To r)
'ReDim arrOperator(1 To r)
'
'' Need to add errors if cannot find table or table columns
'
'For i = 1 To r
'    arrField(i) = tbl.DataBodyRange(i, tbl.ListColumns("Field").Range.Column)
'    arrValue(i) = tbl.DataBodyRange(i, tbl.ListColumns("Value").Range.Column)
'    arrOperator(i) = tbl.DataBodyRange(i, tbl.ListColumns("Operator").Range.Column)
'Next i
'
'For Each pt In ws.PivotTables
'    For Each pf In pt.PivotFields
'        For i = 1 To r
'            If pf.name = arrField(i) Then
'                pf.ClearAllFilters
'                Application.StatusBar = strProcedure & " // " & "PivotName: " & pt.name & "; " & "PivotField: " & pf.name
'                For Each pi In pf.PivotItems
'                    pi.Visible = IIf(Evaluate(Chr(34) & pi.Value & Chr(34) & arrOperator(i) & Chr(34) & arrValue(i) & Chr(34)), True, False)
'                Next pi
'            End If
'        Next i
'    Next pf
'Next pt
'
''   Shutdown
'Application.StatusBar = False
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'Application.Calculation = calcState
'
'Exit Sub
'
'errHandler:
'   Debug.Print "There was an error in this sub" ' reference this sub
'   Application.StatusBar = False
'   Application.ScreenUpdating = screenUpdateState
'   Application.EnableEvents = eventsState
'   Application.Calculation = calcState
'
'End Sub





'Sub PivotDetailA_Show()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'    Dim ws As Worksheet
'    Dim pf As PivotField
'    Dim pt As PivotTable
'
'    Set ws = ActiveSheet
'
'    For Each pt In ws.PivotTables
'        With pt.PivotFields("State")
'            .ShowDetail = True
'        End With
'    Next pt
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'
'Sub pvtDetailB_Show()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'    Dim ws As Worksheet
'    Dim pf As PivotField
'    Dim pt As PivotTable
'
'    Set ws = ActiveSheet
'
'    For Each pt In ws.PivotTables
'        With pt.PivotFields("State")
'            .ShowDetail = True
'        End With
'        With pt.PivotFields("Credit Type")
'            .ShowDetail = True
'        End With
'    Next pt
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'
'Sub pvtDetailA_Hide()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'    Dim ws As Worksheet
'    Dim pf As PivotField
'    Dim pt As PivotTable
'
'    Set ws = ActiveSheet
'
'    For Each pt In ws.PivotTables
'        With pt.PivotFields("State")
'            .ShowDetail = False
'        End With
'        With pt.PivotFields("Credit Type")
'            .ShowDetail = False
'        End With
'    Next pt
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'
'Sub pvtDetailB_Hide()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'    Dim ws As Worksheet
'    Dim pf As PivotField
'    Dim pt As PivotTable
'
'    Set ws = ActiveSheet
'
'    For Each pt In ws.PivotTables
'        With pt.PivotFields("Credit Type")
'            .ShowDetail = False
'        End With
'    Next pt
'
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub

'
'
'Sub pvtClearFilters()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'Dim ws As Worksheet
'Dim pf As PivotField
'Dim pt As PivotTable
'
'Set ws = ActiveSheet
'
''For Each ws In ThisWorkbook.Worksheets
'For Each pt In ws.PivotTables
'    pt.ClearAllFilters
'Next pt
''Next ws
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
'
'Sub pvtAddFields()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim wb As Workbook
'Dim ws As Worksheet
'Dim pf As PivotField
'Dim pt As PivotTable
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
''Set wb = ActiveWorkbook
'Set ws = ActiveSheet
'
''For Each ws In ThisWorkbook.Worksheets
'    For Each pt In ws.PivotTables
'        pt.RefreshTable
'        pt.AddFields (["BreakfieldA","BreakfieldB","BreakfieldC"])
'    Next pt
''Next ws
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub




''''''''''''''''' FILTER PIVOTS BY CELL VALUE

'Private Sub Worksheet_Change(ByVal Target As Range)
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim pf As PivotField
'Dim pi As PivotItem
'Dim strField As String
'Dim BkA As Range
'
'Set ws = ActiveSheet
'Set BkA = Range("C3")
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
''BREAK LEVEL A FILTER
'If Target.Address = BkA.Address Then
'Call pvtRefresh
'Call tblFilterMovie
'strField = "MOVIE NAME"
'    If BkA.Value = "ALL MOVIES" Then
'        GoTo errHandler
'    Else
''clears filters
''        For Each ws In ThisWorkbook.Worksheets
'            For Each pt In ws.PivotTables
'                With pt.PivotFields(strField)
'                .ClearAllFilters
''sets filters to new cell value
'                For Each pi In .PivotItems
'                    If pi.Value = BkA.Value Then
'                        pi.Visible = True
'                    Else
'                        pi.Visible = False
'                    End If
'                Next pi
'                End With
'            Next pt
'        'Next ws
'    End If
'End If
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub



''''OLD VERSION (ALL MOVIES BRINGS BACK ALL MOVIES)
'
'Private Sub Worksheet_Change(ByVal Target As Range)
'
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim pf As PivotField
'Dim pi As PivotItem
'Dim strField As String
'Dim BkA As Range
'
'Set ws = ActiveSheet
'Set BkA = Range("C3")
'
'On Error GoTo errHandler
'Application.EnableEvents = False
'Application.ScreenUpdating = False
'
''BREAK LEVEL A FILTER
'If Target.Address = BkA.Address Then
'strField = "MOVIE NAME"
'    If BkA.Value = "ALL MOVIES" Then
''        GoTo errHandler
''        For Each pt In ws.PivotTables
''            pt.ClearAllFilters
''        Next pt
'        For Each pt In ws.PivotTables
'            With pt.PivotFields(strField)
'            .ClearAllFilters
'                For Each pi In .PivotItems
'                    pi.Visible = True
'                Next pi
'            End With
'        Next pt
'    Else
''clears filters
''        For Each ws In ThisWorkbook.Worksheets
'            For Each pt In ws.PivotTables
'                With pt.PivotFields(strField)
'                .ClearAllFilters
''sets filters to new cell value
'                For Each pi In .PivotItems
'                    If pi.Value = BkA.Value Then
'                        pi.Visible = True
'                    Else
'                        pi.Visible = False
'                    End If
'                Next pi
'                End With
'            Next pt
'        'Next ws
'    End If
'End If
'
'errHandler:
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'
'End Sub









'''''    FILTERS PIVOT CONNECTED TO MULTIPLE TABLES / OLAP DESIGN

''   This filters
'Sub pvtFilterA(strSwitch As String)
'
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim pf As PivotField
'Dim pi As PivotItem
'Dim rngVar As Range
'Dim strField As String
'Dim strVal As String
'
'Set ws = ActiveSheet
'
''Set rngVar = GetRange("Archive")
'strField = "Fund Status"
'strVal = "Open"
'
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'
'    For Each pt In ws.PivotTables
'
'        Select Case pt.Name
'
'            Case "pvtOTBReport"                                                         '   1st pivot table in report
'                With pt.PivotFields(strField)
'                    .ClearAllFilters
'                    'If rngVar.Value = "Yes" Then
'                    If strSwitch = "Yes" Then
'                    For Each pi In .PivotItems
'                        If pi.Value = strVal Or pi.Value Like "*blank*" Then
'                            pi.Visible = True
'                        Else
'                            pi.Visible = False
'                        End If
'                    Next pi
'                    End If
'                End With
'
'            Case "pvtOTBDetail"                                                         '   2nd pivot table in report
'                With pt.PivotFields("[tbl_Demand].[Fund Status].[Fund Status]")
'                    .VisibleItemsList = Array("")           'Clears filters
'                    'If rngVar.Value = "Yes" Then
'                    If strSwitch = "Yes" Then
''                    .VisibleItemsList = Array( _
''                        "[tbl_Demand].[Fund Status].&", _
''                        "[tbl_Demand].[Fund Status].&[Open]")
'                    .VisibleItemsList = Array( _
'                        "[tbl_Demand].[Fund Status].&", _
'                        "[tbl_Demand].[Fund Status].&" & "[" & strVal & "]")
'                    End If
'                End With
'
'        End Select
'    Next pt
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'
'End Sub
