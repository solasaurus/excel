Attribute VB_Name = "mdl_Comments"
Option Explicit

Private Sub ExtractComments_WS()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant
Dim ws As Worksheet
Dim wsC As Worksheet
Dim tbl As ListObject
Dim rng As Range
Dim comm As Comment
Dim i As Integer
Dim x As Integer
Dim strName As String

'   Sub Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Set ws = ActiveSheet
If ws.Comments.Count = 0 Then GoTo NoCommentsError

strName = "Comments"
strWSName = strName

i = 0
CheckWorksheetName:
For Each ws In Worksheets
  If ws.Name = strWSName Then
    i = i + 1
    strWSName = strName & i
    GoTo CheckWorksheetName
  End If
Next ws

Set ws = ActiveSheet

Set wsC = Worksheets.Add(after:=ActiveSheet)
    wsC.Name = strWSName
    
wsC.Range("A1").Value = "Comment In"
wsC.Range("B1").Value = "Comment By"
wsC.Range("C1").Value = "Comment"
Set rng = wsC.Range("A1:C1")
Set tbl = wsC.ListObjects.Add(xlSrcRange, rng, , xlYes)

For Each comm In ws.Comments
    tbl.ListRows.Add (1)
    tbl.DataBodyRange(1, 1).Value = comm.Parent.Address
    tbl.DataBodyRange(1, 2).Value = Left(comm.Text, InStr(1, comm.Text, ":") - 1)
    tbl.DataBodyRange(1, 3).Value = Right(comm.Text, Len(comm.Text) - InStr(1, comm.Text, ":"))
Next comm

x = 20
For i = 1 To 3
    wsC.Columns(i).EntireColumn.AutoFit
    If i = 3 Then
        wsC.Columns(i).ColumnWidth = wsC.Columns(i).ColumnWidth + x
    End If
Next i

'tbl.DataBodyRange.Columns.AutoFit
tbl.DataBodyRange.Rows.AutoFit

errHandler:
    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState

Exit Sub

NoCommentsError:
    MsgBox "No comments found on worksheet with the name: " & Chr(34) & ws.Name & Chr(34), vbExclamation, "NO COMMENTS FOUND"
    GoTo errHandler
    
End Sub

Private Sub ExtractComments_WB()

Dim screenUpdateState As Variant
Dim eventsState As Variant
Dim calcState As Variant
Dim wb As Workbook
Dim ws As Worksheet
Dim wsC As Worksheet
Dim tbl As ListObject
Dim rng As Range
Dim comm As Comment
Dim i As Integer
Dim x As Integer
Dim strName As String
Dim strWSName As String

'   Sub Startup
On Error GoTo errHandler
screenUpdateState = Application.ScreenUpdating
eventsState = Application.EnableEvents
calcState = Application.Calculation

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Set wb = ActiveWorkbook

''' Check if comments exist
i = 0
For Each ws In wb.Worksheets
    If Not ws.Comments.Count = 0 Then
        i = 1
        GoTo CommentsExist
    End If
Next ws

If i = 0 Then GoTo NoCommentsError

''' Check if Comments worksheet name already exists
CommentsExist:
strName = "Comments"
strWSName = strName
i = 0
CheckWorksheetName:
For Each ws In Worksheets
  If ws.Name = strWSName Then
    i = i + 1
    strWSName = strName & i
    GoTo CheckWorksheetName
  End If
Next ws

Set wsC = Worksheets.Add(after:=ActiveSheet)
    wsC.Name = strWSName
    
wsC.Range("A1").Value = "Worksheet"
wsC.Range("B1").Value = "Cell"
wsC.Range("C1").Value = "Comment By"
wsC.Range("D1").Value = "Comment"
Set rng = wsC.Range("A1:D1")
Set tbl = wsC.ListObjects.Add(xlSrcRange, rng, , xlYes)

''' Write to comments table
For Each ws In wb.Worksheets
    If Not ws.Comments.Count = 0 Then
        For Each comm In ws.Comments
            tbl.ListRows.Add (1)
            tbl.DataBodyRange(1, 1).Value = ws.Name
            tbl.DataBodyRange(1, 2).Value = comm.Parent.Address
            tbl.DataBodyRange(1, 3).Value = Left(comm.Text, InStr(1, comm.Text, ":") - 1)
            tbl.DataBodyRange(1, 4).Value = Right(comm.Text, Len(comm.Text) - InStr(1, comm.Text, ":"))
        Next comm
    End If
Next ws

''' Table Cleanup
x = 20
For i = 1 To 4
    wsC.Columns(i).EntireColumn.AutoFit
    If i = 4 Then
        tbl.ListColumns(4).DataBodyRange.WrapText = False
        wsC.Columns(i).ColumnWidth = wsC.Columns(i).ColumnWidth + x
    End If
Next i

tbl.DataBodyRange.Rows.AutoFit

errHandler:
    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState

Exit Sub

NoCommentsError:
    MsgBox "No comments found in " & Chr(34) & wb.Name & Chr(34) & ".", vbExclamation, "NO COMMENTS FOUND"
    GoTo errHandler

End Sub

