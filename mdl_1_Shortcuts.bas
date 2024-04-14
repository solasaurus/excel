Attribute VB_Name = "mdl_1_Shortcuts"
Option Explicit


'   This autoruns on workbook open
Sub Auto_Open()
    init_Shortcuts
End Sub

Private Sub init_Shortcuts()
'   ^ = control _
    + = shift _
    % = alt

'   FORMATTING
    Application.OnKey Key:="^+{x}", procedure:="fmt_QuickNumber"
    Application.OnKey Key:="^+{!}", procedure:="fmt_GenText"
    Application.OnKey Key:="^+{@}", procedure:="fmt_Date"
    Application.OnKey Key:="^+{#}", procedure:="fmt_Number"
    Application.OnKey Key:="^+{$}", procedure:="fmt_Acct"
    Application.OnKey Key:="^+{%}", procedure:="fmt_Percent"
    Application.OnKey Key:="^+{.}", procedure:="fmt_IncrDecimal"
    Application.OnKey Key:="^+{,}", procedure:="fmt_DecrDecimal"
    Application.OnKey Key:="%+{.}", procedure:="fmt_IncrFontSize"
    Application.OnKey Key:="%+{,}", procedure:="fmt_DecrFontSize"

    Application.OnKey Key:="^+{c}", procedure:="fmt_Fontcolor"
    Application.OnKey Key:="^+{d}", procedure:="fmt_Fillcolor"
    Application.OnKey Key:="^+{i}", procedure:="fmt_Input"
    
    Application.OnKey Key:="^+{m}", procedure:="fmt_HAlignment"


'   ACTIONS
    Application.OnKey Key:="^+{v}", procedure:="act_PasteValues"
    Application.OnKey Key:="^+{f}", procedure:="act_PasteFormats"
    Application.OnKey Key:="^+{a}", procedure:="act_AutoFitColumn"
    Application.OnKey Key:="^+{g}", procedure:="act_GroupColumns"
    Application.OnKey Key:="^+{u}", procedure:="act_UngroupColumns"
    
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FORMATTING

Sub fmt_QuickNumber()

Dim a, b, c, d, e, f, g As String

a = "General"
'b = "#,##0"
b = "_(#,##0_);_((#,##0);_( - ??_);_(@_)"
c = "_($* #,##0_);_($* (#,##0);_($* - ??_);_(@_)"
d = "_(* #,##0.0%_);_(* -#,##0.0%_);_(* #,##0.0%_);_(@_)"
e = "m/d/yyyy"
f = "mmm-yy"
g = "_(* #,##0.000_);_(* (#,##0.000);" & Chr(34) & "Check" & Chr(34) & ";" & Chr(34) & "ERROR" & Chr(34)

Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case b
        Selection.NumberFormat = c
    Case c
        Selection.NumberFormat = d
    Case d
        Selection.NumberFormat = e
    Case e
        Selection.NumberFormat = f
    Case f
        Selection.NumberFormat = g
    Case Else
        Selection.NumberFormat = a
End Select

End Sub

Sub fmt_GenText()

Dim a, b, c, d, e, f As String

a = "General"
b = "@"

Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case Else
        Selection.NumberFormat = a
End Select

End Sub

Sub fmt_Date()

Dim a, b, c, d, e, f As String

a = "m/d/yyyy"          ' 3/1/2016
b = "mmm-yy;@"          ' Mar-16
c = "mmmm yyyy;@"       ' March 2016
d = "mmmm d, yyyy;@"    ' March 1, 2016
e = "yyyy;@"            ' 2016

Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case b
        Selection.NumberFormat = c
    Case c
        Selection.NumberFormat = d
    Case d
        Selection.NumberFormat = e
    Case Else
        Selection.NumberFormat = a
End Select

End Sub


Private Sub fmt_HAlignment()

Dim a, b, c, d, e, f As String

a = xlHAlignCenter
b = xlHAlignCenterAcrossSelection
c = xlHAlignLeft
d = xlHAlignRight
e = xlHAlignGeneral


Select Case ActiveCell.NumberFormat
    Case a
        Selection.HorizontalAlignment = b
    Case b
        Selection.HorizontalAlignment = c
    Case c
        Selection.HorizontalAlignment = d
    Case d
        Selection.HorizontalAlignment = e
    Case Else
        Selection.HorizontalAlignment = a
End Select

End Sub

Private Sub fmt_Number()

Dim a, b, c, d, e, f As String

a = "#,##0"
b = "#,##0_);(#,##0)"
c = "_(* #,##0_);_(* (#,##0);_(*  - ??_);_(@_)"
'c = "#,##0_);[Red](#,##0)"
'd = "_(* #,##0_);_(* (#,##0);_(*  - ??_);_(@_)"

Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case b
        Selection.NumberFormat = c
    Case Else
        Selection.NumberFormat = a
End Select

End Sub

Private Sub fmt_Percent()
Dim a, b As String
a = "_(* #,##0.0%_);_(* (#,##0.0%);_(* #,##0.0%_);_(@_)"
b = "#,##0.00"
    
Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case Else
        Selection.NumberFormat = a
End Select

End Sub


Private Sub fmt_Acct()

Dim a, b, c, d, e, f As String

a = "$#,##0"
b = "_($* #,##0_);_($* (#,##0)"
c = "_($* #,##0_);_($* (#,##0.00);_($* - ??_);_(@_)"

Select Case ActiveCell.NumberFormat
    Case a
        Selection.NumberFormat = b
    Case b
        Selection.NumberFormat = c
'    Case d
'        Selection.NumberFormat = e
    Case Else
        Selection.NumberFormat = a
End Select

End Sub

Sub fmt_IncrDecimal()
    Application.CommandBars.FindControl(ID:=398).Execute
End Sub
Sub fmt_DecrDecimal()
    Application.CommandBars.FindControl(ID:=399).Execute
End Sub


Sub fmt_IncrFontSize()
Selection.Font.Size = Selection.Font.Size + 1
End Sub

Sub fmt_DecrFontSize()
Selection.Font.Size = Selection.Font.Size - 1
End Sub

''' COLORS

Sub fmt_Fontcolor()

Dim a, b, c, d As String

a = 1   'black
b = 5   'blue
c = 10  'green
d = 3   'red

Select Case ActiveCell.Font.ColorIndex
    Case a
        Selection.Font.ColorIndex = b
    Case b
        Selection.Font.ColorIndex = c
    Case c
        Selection.Font.ColorIndex = d
    Case Else
        Selection.Font.ColorIndex = a
End Select

End Sub

Sub fmt_Fillcolor()

Dim a, b, c, d, e, f As String

a = -4142   'none
b = 6       'yellow
c = 44      'orange
d = 3       'red
e = 4       'green
f = 19      'yellow input

Select Case ActiveCell.Interior.ColorIndex
    Case a
        Selection.Interior.ColorIndex = b
    Case b
        Selection.Interior.ColorIndex = c
    Case c
        Selection.Interior.ColorIndex = d
    Case d
        Selection.Interior.ColorIndex = e
    Case e
        Selection.Interior.ColorIndex = f
    Case Else
        Selection.Interior.ColorIndex = a
End Select

End Sub


Sub fmt_Input()

Dim a, b, c, d, e, f As String

a = -4142   'none
b = 19

Select Case ActiveCell.Interior.ColorIndex
    Case a
        Selection.Interior.ColorIndex = b
        Selection.Font.ColorIndex = 5
    Case Else
        Selection.Interior.ColorIndex = a
        Selection.Font.ColorIndex = 1
End Select

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
''' ACTIONS

Private Sub act_PasteValues()
If Not Application.CutCopyMode = 0 Then
    Selection.PasteSpecial Paste:=xlValues
End If
End Sub

Private Sub act_PasteFormats()
If Not Application.CutCopyMode = 0 Then
    Selection.PasteSpecial Paste:=xlFormats
End If
End Sub

Private Sub act_AutoFill()
    
    Dim rngVar As Range
    Dim lastRow As Long
    On Error GoTo Shutdown
    
    Set rngVar = Selection
    
    ' If selection is in the first blank cell below the range that should be autofilled then adjust range
    If rngVar = "" Then
        ' If series is only 1 row long then
        If rngVar.Offset(-2, 0) = "" Then
            Set rngVar = rngVar.Offset(-1, 0)
        Else
            Set rngVar = rngVar.Offset(rngVar.Offset(-1, 0).End(xlUp).Row - rngVar.Row, 0)
            Set rngVar = rngVar.Resize(rngVar.End(xlDown).Row - rngVar.Row + 1, 1)
        End If
    End If
    
    ' Determine which adjacent column to autofill off of (use the minimum row length)
    lastRow = WorksheetFunction.Min(rngVar.Offset(0, 1).End(xlDown).Row, rngVar.Offset(0, -1).End(xlDown).Row)
    rngVar.AutoFill Destination:=rngVar.Resize(lastRow - rngVar.Row + 1, 1)
    
Shutdown:

End Sub

Private Sub act_AutoFitColumn()

    Dim rngVar As Range
    Set rngVar = Selection
    rngVar.Columns.AutoFit

End Sub

Private Sub act_GroupColumns()
    Dim selectedRange As Range
    Dim column As Range
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of columns to ungroup."
        Exit Sub
    End If
    
    ' Get the selected range
    Set selectedRange = Selection
    
    ' Loop through each column in the selected range and ungroup
    For Each column In selectedRange.Columns
        column.Group
    Next column
End Sub

Private Sub act_UngroupColumns()
    Dim selectedRange As Range
    Dim column As Range
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of columns to ungroup."
        Exit Sub
    End If
    
    ' Get the selected range
    Set selectedRange = Selection
    
    ' Loop through each column in the selected range and ungroup
    For Each column In selectedRange.Columns
        column.Ungroup
    Next column
End Sub
