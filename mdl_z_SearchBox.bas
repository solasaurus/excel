Attribute VB_Name = "mdl_z_SearchBox"
Option Explicit

''  Add ability for OR search term

'Sub SearchBox()
'Dim ws As Worksheet
'Dim tbl As ListObject
'Dim DataRange As Range
'
'Dim myButton As OptionButton
'Dim SearchString As String
'Dim ButtonName As String
'
'Dim myField As Long
'
'Dim mySearch As Variant
'Dim Box As Range
'
'
''Load Sheet into A Variable
'Set ws = ActiveSheet
'Set tbl = ws.ListObjects("ACQPIPELINE")
'Set DataRange = tbl.Range
'Set Box = Range("A9")
'
''Unfilter Data (if necessary)
'  On Error Resume Next
'    ws.ShowAllData
'    Call ClearSearch
'  On Error GoTo 0
'
''Retrieve User's Search Input
'  mySearch = Box.Value
'
''Determine if user is searching for number or text (currently doesn't matter, but keeping if statement for future
'  If IsNumeric(mySearch) = True Then
'    'SearchString = "=" & mySearch 'original method
'    SearchString = "=*" & mySearch & "*" 'new method that fixes issue when searching for numbers
'  Else
'    SearchString = "=*" & mySearch & "*"
'  End If
'
''Loop Through Option Buttons
'  For Each myButton In ws.OptionButtons
'    If myButton.Value = 1 Then
'      ButtonName = myButton.Text
'      Exit For
'    End If
'  Next myButton
'
''Determine Filter Field
'  On Error GoTo HeadingNotFound
'    myField = Application.WorksheetFunction.Match(ButtonName, DataRange.Rows(1), 0)
'  On Error GoTo 0
'
''Filter Data
'  DataRange.AutoFilter _
'    Field:=myField, _
'    Criteria1:=SearchString, _
'    Operator:=xlAnd
'
''Clear Search Field
'Box.Value = ""
'Box.Select
'ActiveWindow.ScrollRow = tbl.HeaderRowRange.Row             ' Sets view to the top of the table
'
'Call HideAutoFilterDropdowns
'
'Exit Sub
'
''ERROR HANDLERS
'HeadingNotFound:
'  MsgBox "The column heading [" & ButtonName & "] was not found in cells " & DataRange.Rows(1).Address & ". " & _
'    vbNewLine & "Please check for possible typos.", vbCritical, "Header Name Not Found!"
'
'End Sub
'
'
'Sub ClearSearch()
'Dim screenUpdateState As Variant
'Dim eventsState As Variant
'Dim calcState As Variant
'Dim ws As Worksheet
'
'''' Startup
'On Error GoTo errHandler
'screenUpdateState = Application.ScreenUpdating
'eventsState = Application.EnableEvents
'calcState = Application.Calculation
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'Application.Calculation = xlCalculationManual
'
'Set ws = ActiveSheet
'With ws.ListObjects("ACQPIPELINE")
'    .Range.AutoFilter
'    .ShowAutoFilter = True
'End With
'
'Call HideAutoFilterDropdowns
'
'errHandler:
'Application.ScreenUpdating = screenUpdateState
'Application.EnableEvents = eventsState
'Application.Calculation = calcState
'End Sub



