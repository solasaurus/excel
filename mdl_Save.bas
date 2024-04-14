Attribute VB_Name = "mdl_Save"
'Option Explicit

'Sub SaveAsPDF()
'
'Dim dskFPath As String
'Dim sveFName As String
'
'    dskFPath = CreateObject("WScript.Shell").specialfolders("Desktop")
'    sveFName = dskFPath & "\" & Range("A2") & Format(Now, " yy mmdd hhmm") & ".pdf"
'
'    Application.ActiveSheet.ExportAsFixedFormat _
'        Type:=xlTypePDF, _
'        Filename:=sveFName, _
'        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, OpenAfterPublish:=True
'
'End Sub

'Sub Save_PDF()
'
'Dim ws As Worksheet
'Dim strYear          As String: strYear = Year(Date) & "_"
'Dim strMonth         As String: strMonth = Month(Date) & "_"
'Dim strDay           As String: strDay = Day(Date) & "_"
'Dim strReport        As String: strReport = "Report"
'Dim strFileName      As String
'
'Application.DisplayAlerts = False
'
'Set ws = ActiveSheet
''     ThisWorkbook.Sheets(Array("Report Sheet", "Data Sheet")).Select
'     ActiveSheet.ExportAsFixedFormat _
'        Type:=xlTypePDF, _
'        Filename:="Your File Location" & strYear & strMonth & strDay & strReport, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'Application.DisplayAlerts = True
'
'  '  Popup Message that the conversion and save is complete
'     MsgBox "File Saved As:" & vbNewLine & strYear & strMonth & strDay & strReport
'
'End Sub
'
'Sub PDFActiveSheet()
'Dim ws As Worksheet
'Dim strPath As String
'Dim myFile As Variant
'Dim strFile As String
'
'On Error GoTo errHandler
'
'Set ws = ActiveSheet
'
''enter name and select folder for file
'' start in current workbook folder
'strFile = Replace(Replace(ws.Name, " ", ""), ".", "_") _
'            & "_" _
'            & Format(Now(), "yyyymmdd\_hhmm") _
'            & ".pdf"
'strFile = ThisWorkbook.Path & "\" & strFile
'
'myFile = Application.GetSaveAsFilename _
'    (InitialFileName:=strFile, _
'        FileFilter:="PDF Files (*.pdf), *.pdf", _
'        Title:="Select Folder and FileName to save")
'
'If myFile <> "False" Then
'    ws.ExportAsFixedFormat _
'        Type:=xlTypePDF, _
'        Filename:=myFile, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    MsgBox "PDF file has been created."
'    ' add yes/no to either open or not
'End If
'
'exitHandler:
'    Exit Sub
'errHandler:
'    MsgBox "Could not create PDF file"
'    Resume exitHandler
'End Sub



'   Extends the bottom of the print area to the maxRow of a particular pivottable on the worksheet
'Sub SetPrintArea()
'
'Dim ws As Worksheet
'Dim pt As PivotTable
'Dim varRng As Range
'Dim varStr As String
'Dim maxRow As Long
'Dim varLong As Long
'
'Set ws = ActiveSheet
'Set pt = ws.PivotTables("pvt_STPayments")
'
''   Check if AutoSet Print Area variable is chosen
'If Range("D7").Value = "Yes" Then
'
''   Get last row of pivot table
'With pt.TableRange1
'    maxRow = .Cells(.Cells.Count).Row
'End With
'
'With ws
'    Set varRng = Range(.PageSetup.PrintArea)                '   Original print area
'    varLong = maxRow - varRng.Row + 1                       '   Difference between last row of pivot table and first row of current print area = new row length of print area
'    .PageSetup.PrintArea = varRng.Resize(varLong).Address   '   New print area
'End With
'
'End If
'
'End Sub
